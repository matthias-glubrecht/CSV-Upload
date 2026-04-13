import CsvUploadService from './CsvUploadService';
import TaxonomyProcessor from './TaxonomyProcessor';
import {
  ICsvData,
  IFieldMapping,
  IImportProgress,
  IFieldErrorInfo,
  FieldErrorDecision
} from '../models';
import { resolveDefaultValue, extractErrorMessage } from '../utils/ImportHelpers';
import * as strings from 'UploadCsvWebPartStrings';

const LOG: string = '[CsvUpload]';

/**
 * Callback interface that decouples the import engine from React state.
 * The component implements these to update the UI during import.
 */
export interface IImportCallbacks {
  /** Called whenever import progress changes (row completed, error, etc.) */
  onProgress(progress: IImportProgress): void;
  /**
   * Called when a field value cannot be converted. The import pauses
   * until the returned Promise resolves with the user's decision.
   */
  onFieldError(error: IFieldErrorInfo): Promise<FieldErrorDecision>;
}

/**
 * Orchestrates the CSV-to-SharePoint import process.
 *
 * Processes rows sequentially, converts field values, handles
 * create/update (upsert via key column), and delegates taxonomy
 * field processing to TaxonomyProcessor.
 */
export default class ImportEngine {

  private _service: CsvUploadService;
  private _taxonomyProcessor: TaxonomyProcessor;

  constructor(service: CsvUploadService) {
    this._service = service;
    this._taxonomyProcessor = new TaxonomyProcessor(service);
  }

  /**
   * Run the full import of CSV data into a SharePoint list.
   *
   * @param csvData   Parsed CSV (headers + rows)
   * @param mappings  Active field mappings (only mapped or defaulted fields)
   * @param webUrl    Target web URL
   * @param listId    Target list ID
   * @param callbacks UI callbacks for progress updates and error prompts
   * @returns The final progress state
   */
  public run(
    csvData: ICsvData,
    mappings: IFieldMapping[],
    webUrl: string,
    listId: string,
    callbacks: IImportCallbacks
  ): Promise<IImportProgress> {
    const activeMappings: IFieldMapping[] = mappings.filter(
      (m: IFieldMapping) =>
        m.csvColumn !== undefined ||
        (m.defaultValue !== undefined && m.defaultValue !== '')
    );
    const keyMapping: IFieldMapping | undefined =
      activeMappings.filter((m: IFieldMapping) => m.isKeyColumn)[0];

    // Separate taxonomy mappings (processed after item create/update)
    const taxonomyMappings: IFieldMapping[] = activeMappings.filter(
      (m: IFieldMapping) =>
        m.listField.fieldType === 'TaxonomyFieldType' ||
        m.listField.fieldType === 'TaxonomyFieldTypeMulti'
    );
    const nonTaxonomyMappings: IFieldMapping[] = activeMappings.filter(
      (m: IFieldMapping) =>
        m.listField.fieldType !== 'TaxonomyFieldType' &&
        m.listField.fieldType !== 'TaxonomyFieldTypeMulti'
    );

    console.log(LOG, 'ImportEngine.run \u2014 activeMappings:', activeMappings.length,
      'taxonomyMappings:', taxonomyMappings.length,
      taxonomyMappings.map((m: IFieldMapping) => ({
        field: m.listField.internalName,
        type: m.listField.fieldType,
        csvColumn: m.csvColumn,
        termSetId: m.listField.termSetId
      })),
      'nonTaxonomyMappings:', nonTaxonomyMappings.length
    );

    const progress: IImportProgress = {
      total: csvData.rows.length,
      current: 0,
      created: 0,
      updated: 0,
      errors: 0,
      errorMessages: [],
      status: 'running'
    };
    callbacks.onProgress(progress);

    // Step 1: If we have a key column, look up existing items first
    const existingItemsPromise: Promise<{ [key: string]: number }> = keyMapping
      ? this._getExistingItemsMap(webUrl, listId, keyMapping, csvData)
      : Promise.resolve({});

    return existingItemsPromise.then(
      (existingItems: { [key: string]: number }) => {
        // Step 2: Process rows sequentially
        return this._processRows(
          csvData.rows, 0, webUrl, listId,
          nonTaxonomyMappings, taxonomyMappings,
          csvData.headers, keyMapping, existingItems,
          progress, callbacks
        );
      }
    ).then(() => {
      progress.status = 'completed';
      callbacks.onProgress({ ...progress });
      return progress;
    // tslint:disable-next-line:no-any
    }).catch((error: Error) => {
      progress.status = 'error';
      progress.errorMessages.push(
        strings.ImportAborted + (error.message || error)
      );
      callbacks.onProgress({ ...progress });
      return progress;
    });
  }

  // ─── Existing Items Lookup ─────────────────────────────────────

  private _getExistingItemsMap(
    webUrl: string,
    listId: string,
    keyMapping: IFieldMapping,
    csvData: ICsvData
  ): Promise<{ [key: string]: number }> {
    const keyColumnIndex: number = csvData.headers.indexOf(keyMapping.csvColumn);
    if (keyColumnIndex < 0) {
      return Promise.resolve({});
    }

    const keyValues: string[] = csvData.rows.map((row: string[]) => {
      return keyColumnIndex < row.length ? row[keyColumnIndex] : '';
    }).filter((v: string) => v.length > 0);

    return this._service.getExistingItems(
      webUrl, listId, keyMapping.listField.internalName, keyValues
    );
  }

  // ─── Row Processing ───────────────────────────────────────────

  private _processRows(
    rows: string[][],
    index: number,
    webUrl: string,
    listId: string,
    nonTaxMappings: IFieldMapping[],
    taxMappings: IFieldMapping[],
    headers: string[],
    keyMapping: IFieldMapping | undefined,
    existingItems: { [key: string]: number },
    progress: IImportProgress,
    callbacks: IImportCallbacks
  ): Promise<void> {
    if (index >= rows.length) {
      return Promise.resolve();
    }

    const row: string[] = rows[index];

    return this._processRow(
      row, webUrl, listId, nonTaxMappings, taxMappings,
      headers, keyMapping, existingItems, progress, callbacks
    ).then(() => {
      progress.current = index + 1;
      callbacks.onProgress({ ...progress });
      return this._processRows(
        rows, index + 1, webUrl, listId, nonTaxMappings,
        taxMappings, headers, keyMapping, existingItems,
        progress, callbacks
      );
    });
  }

  private _processRow(
    row: string[],
    webUrl: string,
    listId: string,
    nonTaxMappings: IFieldMapping[],
    taxMappings: IFieldMapping[],
    headers: string[],
    keyMapping: IFieldMapping | undefined,
    existingItems: { [key: string]: number },
    progress: IImportProgress,
    callbacks: IImportCallbacks
  ): Promise<void> {
    const service: CsvUploadService = this._service;
    const rowNum: number = progress.current + 1;

    // tslint:disable-next-line:no-any
    type FieldResult = { fieldName: string; value: any };
    // tslint:disable-next-line:no-any
    const fieldValues: { [key: string]: any } = {};
    let skipRow: boolean = false;

    /**
     * Process field mappings sequentially so we can catch each error
     * individually and show the correction dialog.
     */
    const processField: (idx: number) => Promise<void> = (idx: number) => {
      if (idx >= nonTaxMappings.length || skipRow) {
        return Promise.resolve();
      }

      const mapping: IFieldMapping = nonTaxMappings[idx];
      const colIndex: number = headers.indexOf(mapping.csvColumn);
      let csvValue: string = colIndex >= 0 && colIndex < row.length
        ? row[colIndex] : '';
      if ((!csvValue || csvValue === '') && mapping.defaultValue) {
        csvValue = resolveDefaultValue(
          mapping.defaultValue, mapping.listField.fieldType
        );
      }

      return this._convertFieldWithRetry(
        webUrl, mapping, csvValue, rowNum, callbacks
      ).then((result: FieldResult | undefined) => {
        if (result && (result as { skipRow?: boolean }).skipRow) {
          skipRow = true;
          return;
        }
        if (result) {
          fieldValues[result.fieldName] = result.value;
        }
        return processField(idx + 1);
      });
    };

    return processField(0).then(() => {
      if (skipRow) {
        return;
      }

      // Determine if this is an update or create
      let existingItemId: number | undefined = undefined;
      if (keyMapping) {
        const keyColIndex: number = headers.indexOf(keyMapping.csvColumn);
        const keyVal: string = keyColIndex >= 0 && keyColIndex < row.length
          ? row[keyColIndex] : '';
        if (keyVal && existingItems[keyVal] !== undefined) {
          existingItemId = existingItems[keyVal];
        }
      }

      // tslint:disable-next-line:no-any
      const itemPromise: Promise<{ Id: number }> = existingItemId !== undefined
        ? service.updateItem(webUrl, listId, existingItemId, fieldValues)
            .then(() => ({ Id: existingItemId }))
        : service.createItem(webUrl, listId, fieldValues)
            .then((created: { Id: number }) => {
              progress.created++;
              return created;
            });

      if (existingItemId !== undefined) {
        itemPromise.then(() => { progress.updated++; });
      }

      return itemPromise.then((item: { Id: number }) => {
        // Handle taxonomy fields
        if (taxMappings.length > 0 && item && item.Id) {
          console.log(LOG, '_processRow \u2014 processing taxonomy fields for itemId:',
            item.Id, 'taxMappings:', taxMappings.length);
          return this._taxonomyProcessor.processTaxonomyFields(
            row, webUrl, listId, item.Id, taxMappings,
            headers, rowNum, callbacks.onFieldError
          );
        } else {
          console.log(LOG, '_processRow \u2014 no taxonomy processing.',
            'taxMappings.length:', taxMappings.length,
            'item:', item ? 'Id=' + item.Id : '(undefined)');
        }
      });
    // tslint:disable-next-line:no-any
    }).catch((error: Error) => {
      // Let abort signal propagate — do not swallow it
      if (error && error.message === '__ABORT_IMPORT__') {
        throw error;
      }
      progress.errors++;
      const friendlyMsg: string = extractErrorMessage(error);
      progress.errorMessages.push(
        strings.ErrorRowPrefix.replace('{0}', String(rowNum)) + friendlyMsg
      );
      if (progress.errorMessages.length > 100) {
        progress.errorMessages = progress.errorMessages.slice(0, 100);
        progress.errorMessages.push(strings.ErrorTruncated);
      }
    });
  }

  // ─── Field Conversion with Retry ──────────────────────────────

  /**
   * Attempt to convert a field value; if conversion fails, show the
   * error dialog and let the user correct, skip the field, or skip the row.
   */
  private _convertFieldWithRetry(
    webUrl: string,
    mapping: IFieldMapping,
    csvValue: string,
    rowNum: number,
    callbacks: IImportCallbacks
    // tslint:disable-next-line:no-any
  ): Promise<{ fieldName: string; value: any; skipRow?: boolean } | undefined> {
    const service: CsvUploadService = this._service;

    return service.convertFieldValue(
      webUrl, mapping.listField, csvValue, mapping.allowFillIn
    // tslint:disable-next-line:no-any
    ).catch((error: any) => {
      const friendlyMsg: string = extractErrorMessage(error);
      const errorInfo: IFieldErrorInfo = {
        rowNumber: rowNum,
        fieldDisplayName: mapping.listField.displayName || mapping.listField.internalName,
        fieldInternalName: mapping.listField.internalName,
        fieldType: mapping.listField.fieldType,
        csvValue: csvValue,
        errorMessage: friendlyMsg
      };

      return callbacks.onFieldError(errorInfo).then(
        (decision: FieldErrorDecision) => {
          if (decision.action === 'skip-row') {
            return { fieldName: '', value: undefined, skipRow: true };
          }
          if (decision.action === 'abort-import') {
            throw new Error('__ABORT_IMPORT__');
          }
          if (decision.action === 'skip-field') {
            return undefined;
          }
          // 'use-value' — retry with the corrected value
          return this._convertFieldWithRetry(
            webUrl, mapping, decision.newValue, rowNum, callbacks
          );
        }
      );
    });
  }
}
