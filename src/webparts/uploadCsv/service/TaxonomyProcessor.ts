import CsvUploadService from './CsvUploadService';
import { IFieldMapping, IFieldErrorInfo, FieldErrorDecision } from '../models';
import { parseTaxonomyExportValue, resolveDefaultValue } from '../utils/ImportHelpers';
import * as strings from 'UploadCsvWebPartStrings';

const LOG: string = '[CsvUpload]';

/** Result returned by taxonomy resolution methods */
interface ITaxonomyResult {
  wssId: number;
  label: string;
  termGuid: string;
}

/**
 * Handles taxonomy (managed metadata) field processing during CSV import.
 * Resolves taxonomy labels through a two-tier strategy:
 *   1. TaxonomyHiddenList lookup (fast, cached terms)
 *   2. Term store lookup (fallback for unused terms)
 */
export default class TaxonomyProcessor {

  private _service: CsvUploadService;

  constructor(service: CsvUploadService) {
    this._service = service;
  }

  /**
   * Process all taxonomy field mappings for a single list item.
   * Called after the item has been created/updated, because SharePoint
   * requires the item to exist before taxonomy values can be set.
   *
   * @param row          The CSV row values
   * @param webUrl       Target web URL
   * @param listId       Target list ID
   * @param itemId       The ID of the already-created/updated item
   * @param taxMappings  Only taxonomy field mappings
   * @param headers      CSV header names (to resolve column indices)
   * @param rowNum       1-based row number for error messages
   * @param onFieldError Callback to pause and prompt the user on error
   */
  public processTaxonomyFields(
    row: string[],
    webUrl: string,
    listId: string,
    itemId: number,
    taxMappings: IFieldMapping[],
    headers: string[],
    rowNum: number,
    onFieldError: (error: IFieldErrorInfo) => Promise<FieldErrorDecision>
  ): Promise<void> {
    const processMapping: (idx: number) => Promise<void> = (idx: number) => {
      if (idx >= taxMappings.length) {
        return Promise.resolve();
      }

      const mapping: IFieldMapping = taxMappings[idx];
      const colIndex: number = headers.indexOf(mapping.csvColumn);
      let csvValue: string = colIndex >= 0 && colIndex < row.length
        ? row[colIndex] : '';
      // If cell is empty, fall back to the configured default value
      if ((!csvValue || csvValue === '') && mapping.defaultValue) {
        csvValue = resolveDefaultValue(
          mapping.defaultValue, mapping.listField.fieldType
        );
      }

      console.log(LOG, 'processTaxonomyFields \u2014 idx:', idx,
        'field:', mapping.listField.internalName,
        'type:', mapping.listField.fieldType,
        'csvColumn:', mapping.csvColumn,
        'colIndex:', colIndex,
        'csvValue:', JSON.stringify(csvValue),
        'termSetId:', mapping.listField.termSetId);

      if (!csvValue) {
        console.log(LOG, 'processTaxonomyFields \u2014 skipping (empty value)');
        return processMapping(idx + 1);
      }

      if (mapping.listField.fieldType === 'TaxonomyFieldTypeMulti') {
        return this._processMultiTaxonomy(
          webUrl, listId, itemId, mapping, csvValue, rowNum, onFieldError
        ).then(() => processMapping(idx + 1));
      }

      return this._processSingleTaxonomy(
        webUrl, listId, itemId, mapping, csvValue, rowNum, onFieldError
      ).then(() => processMapping(idx + 1));
    };

    return processMapping(0);
  }

  // ─── Multi-Value Taxonomy ──────────────────────────────────────

  private _processMultiTaxonomy(
    webUrl: string,
    listId: string,
    itemId: number,
    mapping: IFieldMapping,
    csvValue: string,
    rowNum: number,
    onFieldError: (error: IFieldErrorInfo) => Promise<FieldErrorDecision>
  ): Promise<void> {
    const service: CsvUploadService = this._service;
    const labels: string[] = parseTaxonomyExportValue(csvValue);
    const tsId: string = mapping.listField.termSetId || '';
    const hiddenField: string = mapping.listField.taxonomyHiddenFieldName
      || mapping.listField.internalName;

    const termPromises: Promise<ITaxonomyResult | undefined>[] =
      labels.map((lbl: string) =>
        service.resolveTaxonomyValue(webUrl, tsId, lbl)
      );

    return Promise.all(termPromises).then(
      (results: (ITaxonomyResult | undefined)[]) => {
        const validResults: ITaxonomyResult[] = results.filter(
          (r: ITaxonomyResult | undefined) => r !== undefined
        ) as ITaxonomyResult[];
        // Collect labels that were NOT found in TaxonomyHiddenList
        const unresolvedLabels: string[] = labels.filter(
          (_lbl: string, i: number) => results[i] === undefined
        );
        console.log(LOG, 'processTaxonomyFields multi \u2014 labels:', labels,
          'resolved:', validResults.length, 'of', labels.length, validResults,
          'unresolved:', unresolvedLabels);

        // Resolve unresolved labels from the term store
        const storePromises: Promise<ITaxonomyResult | undefined>[] =
          unresolvedLabels.map((ul: string) =>
            service.resolveTermFromStore(tsId, ul, mapping.listField.sspId)
          );
        return Promise.all(storePromises).then(
          (storeResults: (ITaxonomyResult | undefined)[]) => {
            const storeResolved: ITaxonomyResult[] = storeResults.filter(
              (r: ITaxonomyResult | undefined) => r !== undefined
            ) as ITaxonomyResult[];
            // Collect labels still unresolved after term store lookup
            const stillUnresolved: string[] = unresolvedLabels.filter(
              (_lbl: string, i: number) => storeResults[i] === undefined
            );
            if (stillUnresolved.length > 0) {
              const fieldLabel: string = mapping.listField.displayName
                || mapping.listField.internalName;
              console.warn(LOG,
                'processTaxonomyFields multi \u2014 terms not found:',
                stillUnresolved, 'for field:', fieldLabel);
              throw new Error(
                strings.ErrorTaxonomyTermNotFound
                  .replace('{0}', stillUnresolved.join(', '))
                  .replace('{1}', fieldLabel)
              );
            }
            const allValues: ITaxonomyResult[] = validResults.concat(storeResolved);
            if (allValues.length > 0) {
              return service.setTaxonomyMultiFieldValue(
                webUrl, listId, itemId,
                mapping.listField.internalName, hiddenField,
                allValues
              );
            }
          }
        );
      }
    )
    // tslint:disable-next-line:no-any
    .catch((err: any) => {
      return this._handleTaxonomyError(
        err, mapping, csvValue, rowNum, onFieldError
      );
    });
  }

  // ─── Single-Value Taxonomy ─────────────────────────────────────

  private _processSingleTaxonomy(
    webUrl: string,
    listId: string,
    itemId: number,
    mapping: IFieldMapping,
    csvValue: string,
    rowNum: number,
    onFieldError: (error: IFieldErrorInfo) => Promise<FieldErrorDecision>
  ): Promise<void> {
    const service: CsvUploadService = this._service;

    // Parse export format and take first label
    const parsedLabels: string[] = parseTaxonomyExportValue(csvValue);
    if (parsedLabels.length > 1) {
      console.warn(LOG, 'processTaxonomyFields single \u2014 multiple values found,',
        'using first label only:', JSON.stringify(parsedLabels[0]),
        'from:', JSON.stringify(csvValue));
    }
    const label: string = parsedLabels.length > 0 ? parsedLabels[0] : csvValue;
    const termSetId: string = mapping.listField.termSetId || '';
    const taxHiddenField: string = mapping.listField.taxonomyHiddenFieldName
      || mapping.listField.internalName;

    return service.resolveTaxonomyValue(webUrl, termSetId, label)
      .then((result: ITaxonomyResult | undefined) => {
        console.log(LOG, 'processTaxonomyFields single \u2014 csvValue:',
          JSON.stringify(label),
          'resolved:', result ? JSON.stringify(result) : '(not found, will search term store)');
        if (result) {
          return service.setTaxonomyFieldValue(
            webUrl, listId, itemId,
            mapping.listField.internalName, taxHiddenField,
            result.label, result.termGuid
          );
        }
        // Term not in TaxonomyHiddenList — search term store for the GUID
        return service.resolveTermFromStore(
          termSetId, label, mapping.listField.sspId
        ).then((storeResult: ITaxonomyResult | undefined) => {
          if (storeResult) {
            console.log(LOG, 'processTaxonomyFields single \u2014 resolved from STORE:',
              JSON.stringify(storeResult),
              'will call setTaxonomyFieldValue with:',
              'itemId:', itemId,
              'field:', mapping.listField.internalName,
              'hiddenField:', taxHiddenField,
              'label:', storeResult.label,
              'termGuid:', storeResult.termGuid);
            return service.setTaxonomyFieldValue(
              webUrl, listId, itemId,
              mapping.listField.internalName, taxHiddenField,
              storeResult.label, storeResult.termGuid
            );
          }
          console.warn(LOG, 'processTaxonomyFields \u2014 term not found anywhere:',
            JSON.stringify(label),
            'termSetId:', termSetId,
            'sspId:', mapping.listField.sspId);
          throw new Error(
            strings.ErrorTaxonomyTermNotFound
              .replace('{0}', label)
              .replace('{1}', mapping.listField.displayName || mapping.listField.internalName)
          );
        });
      })
      // tslint:disable-next-line:no-any
      .catch((err: any) => {
        return this._handleTaxonomyError(
          err, mapping, csvValue, rowNum, onFieldError
        );
      });
  }

  // ─── Shared Error Handler ──────────────────────────────────────

  /**
   * Translate SharePoint taxonomy errors into a friendly message
   * and show the field error dialog. If the user chooses 'skip-row',
   * a special __SKIP_ROW__ error is thrown to abort the row.
   */
  // tslint:disable-next-line:no-any
  private _handleTaxonomyError(
    err: Error,
    mapping: IFieldMapping,
    csvValue: string,
    rowNum: number,
    onFieldError: (error: IFieldErrorInfo) => Promise<FieldErrorDecision>
  ): Promise<void> {
    const msg: string = err && err.message ? err.message : String(err);
    let friendlyMsg: string = msg;
    if (msg.indexOf('SPFieldValueException') >= 0
      || msg.indexOf('Terminologiespeicher') >= 0
      || msg.indexOf('term store') >= 0) {
      friendlyMsg = strings.ErrorTaxonomyTermNotInTermSet
        .replace('{0}', csvValue)
        .replace('{1}', mapping.listField.displayName || mapping.listField.internalName);
    }
    const errorInfo: IFieldErrorInfo = {
      rowNumber: rowNum,
      fieldDisplayName: mapping.listField.displayName || mapping.listField.internalName,
      fieldInternalName: mapping.listField.internalName,
      fieldType: mapping.listField.fieldType,
      csvValue: csvValue,
      errorMessage: friendlyMsg
    };
    return onFieldError(errorInfo).then(
      (decision: FieldErrorDecision) => {
        if (decision.action === 'skip-row') {
          throw new Error('__SKIP_ROW__');
        }
        if (decision.action === 'abort-import') {
          throw new Error('__ABORT_IMPORT__');
        }
        // skip-field or use-value: just continue to next mapping
        // (taxonomy re-resolve with a corrected value is complex;
        // the user can re-import to fix remaining issues)
      }
    );
  }
}
