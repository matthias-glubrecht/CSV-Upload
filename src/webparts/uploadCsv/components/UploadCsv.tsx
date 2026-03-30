import * as React from 'react';
import styles from './UploadCsv.module.scss';
import { IUploadCsvProps } from './IUploadCsvProps';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import {
  ISiteCollection,
  IWeb,
  IListInfo,
  IListField,
  IFieldMapping,
  ICsvData,
  IImportProgress,
  IFieldErrorInfo,
  FieldErrorDecision
} from '../models';
import CsvUploadService from '../service/CsvUploadService';
import { parseCsv, detectDelimiter, decodeCsvBytes } from '../utils/CsvParser';
import SiteCollectionPicker from './SiteCollectionPicker/SiteCollectionPicker';
import WebPicker from './WebPicker/WebPicker';
import ListPicker from './ListPicker/ListPicker';
import CsvDropZone from './CsvDropZone/CsvDropZone';
import MappingTable from './MappingTable/MappingTable';
import ImportProgress from './ImportProgress/ImportProgress';
import FieldErrorDialog from './FieldErrorDialog/FieldErrorDialog';
import * as strings from 'UploadCsvWebPartStrings';

const LOG: string = '[CsvUpload]';

export interface IUploadCsvState {
  selectedSiteCollection: ISiteCollection | undefined;
  selectedWeb: IWeb | undefined;
  selectedList: IListInfo | undefined;
  listFields: IListField[];
  csvData: ICsvData | undefined;
  csvFileName: string | undefined;
  mappings: IFieldMapping[];
  progress: IImportProgress;
  loadingFields: boolean;
  errorMessage: string | undefined;
  /** When set, the FieldErrorDialog is shown and the import is paused */
  fieldError: IFieldErrorInfo | undefined;
}

export default class UploadCsv extends React.Component<IUploadCsvProps, IUploadCsvState> {

  /** Stored resolve function for the field error dialog promise */
  private _fieldErrorResolve: ((decision: FieldErrorDecision) => void) | undefined;

  constructor(props: IUploadCsvProps) {
    super(props);

    const currentSite: ISiteCollection = props.service.getCurrentSiteCollection();
    const currentWeb: IWeb = props.service.getCurrentWeb();

    this.state = {
      selectedSiteCollection: currentSite,
      selectedWeb: currentWeb,
      selectedList: undefined,
      listFields: [],
      csvData: undefined,
      csvFileName: undefined,
      mappings: [],
      progress: {
        total: 0,
        current: 0,
        created: 0,
        updated: 0,
        errors: 0,
        errorMessages: [],
        status: 'idle'
      },
      loadingFields: false,
      errorMessage: undefined,
      fieldError: undefined
    };
  }

  public render(): React.ReactElement<IUploadCsvProps> {
    const {
      selectedSiteCollection,
      selectedWeb,
      selectedList,
      csvData,
      csvFileName,
      mappings,
      progress,
      loadingFields,
      errorMessage
    } = this.state;

    const canImport: boolean = this._canImport();
    const hasKeyColumn: boolean = mappings.some((m: IFieldMapping) => m.isKeyColumn);

    return (
      <div className={styles.uploadCsv}>
        <div className={styles.container}>
          {/* Header */}
          <div className={styles.header}>{strings.WebPartTitle}</div>
          <div className={styles.subHeader}>
            {strings.WebPartSubTitle}
          </div>

          {/* Selectors */}
          <div className={styles.section}>
            <div className={styles.sectionTitle}>{strings.SectionTargetTitle}</div>
            <div className={styles.selectorRow}>
              <div className={styles.selectorColumn}>
                <SiteCollectionPicker
                  service={this.props.service}
                  selectedSiteCollection={selectedSiteCollection}
                  onSiteCollectionChanged={this._onSiteCollectionChanged}
                />
              </div>
              <div className={styles.selectorColumn}>
                {selectedSiteCollection && (
                  <WebPicker
                    service={this.props.service}
                    siteCollectionUrl={selectedSiteCollection.url}
                    selectedWeb={selectedWeb}
                    onWebChanged={this._onWebChanged}
                  />
                )}
              </div>
              <div className={styles.selectorColumn}>
                {selectedWeb && (
                  <ListPicker
                    service={this.props.service}
                    webUrl={selectedWeb.url}
                    selectedList={selectedList}
                    onListChanged={this._onListChanged}
                  />
                )}
              </div>
            </div>
          </div>

          {/* CSV Drop Zone */}
          {selectedList && (
            <div className={styles.section}>
              <div className={styles.sectionTitle}>{strings.SectionUploadTitle}</div>
              {loadingFields ? (
                <Spinner size={SpinnerSize.large} label={strings.LoadingFieldsLabel} />
              ) : (
                <CsvDropZone
                  onFileSelected={this._onFileSelected}
                  fileName={csvFileName}
                />
              )}
            </div>
          )}

          {/* Mapping Table */}
          {csvData && mappings.length > 0 && (
            <div className={styles.section}>
              <div className={styles.sectionTitle}>{strings.SectionMappingTitle}</div>
              <MappingTable
                mappings={mappings}
                csvHeaders={csvData.headers}
                onMappingChanged={this._onMappingChanged}
              />

              {/* Validation messages */}
              {!this._allRequiredFieldsMapped() && (
                <div className={styles.validationMessage}>
                  {strings.ValidationRequiredFields}
                </div>
              )}

              {!hasKeyColumn && (
                <div className={styles.validationMessage}>
                  {strings.ValidationNoKeyColumn}
                </div>
              )}

              {/* Import button */}
              <div className={styles.buttonRow}>
                <PrimaryButton
                  text={strings.ImportButtonText}
                  onClick={this._onImport}
                  disabled={!canImport || progress.status === 'running'}
                  iconProps={{ iconName: 'Upload' }}
                />
                <DefaultButton
                  text={strings.ResetButtonText}
                  onClick={this._onReset}
                  disabled={progress.status === 'running'}
                  iconProps={{ iconName: 'Refresh' }}
                />
              </div>
            </div>
          )}

          {/* Progress */}
          <ImportProgress progress={progress} />

          {/* Error Message */}
          {errorMessage && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={this._clearError}>
              {errorMessage}
            </MessageBar>
          )}

          {/* Field Error Dialog */}
          {this.state.fieldError && (
            <FieldErrorDialog
              error={this.state.fieldError}
              onDecision={this._onFieldErrorDecision}
            />
          )}
        </div>
      </div>
    );
  }

  // ─── Event Handlers ──────────────────────────────────────────────────

  private _onSiteCollectionChanged = (siteCollection: ISiteCollection): void => {
    this.setState({
      selectedSiteCollection: siteCollection,
      selectedWeb: undefined,
      selectedList: undefined,
      listFields: [],
      csvData: undefined,
      csvFileName: undefined,
      mappings: [],
      progress: this._resetProgress()
    });
  }

  private _onWebChanged = (web: IWeb): void => {
    this.setState({
      selectedWeb: web,
      selectedList: undefined,
      listFields: [],
      csvData: undefined,
      csvFileName: undefined,
      mappings: [],
      progress: this._resetProgress()
    });
  }

  private _onListChanged = (list: IListInfo): void => {
    this.setState({
      selectedList: list,
      listFields: [],
      csvData: undefined,
      csvFileName: undefined,
      mappings: [],
      loadingFields: true,
      progress: this._resetProgress()
    });

    const webUrl: string = this.state.selectedWeb.url;
    this.props.service.getListFields(webUrl, list.id).then((fields: IListField[]) => {
      this.setState({
        listFields: fields,
        loadingFields: false
      });
    // tslint:disable-next-line:no-any
    }).catch((error: Error) => {
      this.setState({
        loadingFields: false,
        errorMessage: strings.ErrorLoadingFields + (error.message || error)
      });
    });
  }

  private _onFileSelected = (file: File): void => {
    const reader: FileReader = new FileReader();
    reader.onload = (e: ProgressEvent) => {
      // tslint:disable-next-line:no-any
      const buffer: ArrayBuffer = (e.target as any).result as ArrayBuffer;
      const csvText: string = decodeCsvBytes(buffer);
      const delimiter: string = detectDelimiter(csvText);
      const csvData: ICsvData = parseCsv(csvText, delimiter);

      if (csvData.headers.length === 0) {
        this.setState({ errorMessage: strings.ErrorNoHeaders });
        return;
      }

      const mappings: IFieldMapping[] = this._createInitialMappings(this.state.listFields, csvData.headers);

      this.setState({
        csvData: csvData,
        csvFileName: file.name,
        mappings: mappings,
        progress: this._resetProgress()
      });
    };
    reader.onerror = () => {
      this.setState({ errorMessage: strings.ErrorReadingFile });
    };
    reader.readAsArrayBuffer(file);
  }

  private _onMappingChanged = (updatedMappings: IFieldMapping[]): void => {
    this.setState({ mappings: updatedMappings });
  }

  private _onImport = (): void => {
    this._importData();
  }

  private _onReset = (): void => {
    this.setState({
      csvData: undefined,
      csvFileName: undefined,
      mappings: [],
      progress: this._resetProgress(),
      errorMessage: undefined,
      fieldError: undefined
    });
  }

  private _clearError = (): void => {
    this.setState({ errorMessage: undefined });
  }

  /**
   * Called by the FieldErrorDialog when the user makes a decision.
   * Resolves the stored promise so the import chain can continue.
   */
  private _onFieldErrorDecision = (decision: FieldErrorDecision): void => {
    this.setState({ fieldError: undefined });
    if (this._fieldErrorResolve) {
      const resolve: (d: FieldErrorDecision) => void = this._fieldErrorResolve;
      this._fieldErrorResolve = undefined;
      resolve(decision);
    }
  }

  /**
   * Show the field error dialog and return a Promise that resolves
   * when the user makes a decision. This pauses the import chain.
   */
  private _showFieldErrorDialog(error: IFieldErrorInfo): Promise<FieldErrorDecision> {
    return new Promise<FieldErrorDecision>((resolve: (d: FieldErrorDecision) => void) => {
      this._fieldErrorResolve = resolve;
      this.setState({ fieldError: error });
    });
  }

  // ─── Mapping Logic ───────────────────────────────────────────────────

  private _createInitialMappings(fields: IListField[], csvHeaders: string[]): IFieldMapping[] {
    const lowerHeaders: string[] = csvHeaders.map((h: string) => h.toLowerCase());
    const mappedHeaders: string[] = [];

    const mappings: IFieldMapping[] = fields.map((field: IListField) => {
      let matchedColumn: string | undefined = undefined;

      // First try matching by display name
      const displayNameLower: string = field.displayName.toLowerCase();
      const displayNameIndex: number = lowerHeaders.indexOf(displayNameLower);
      if (displayNameIndex >= 0 && mappedHeaders.indexOf(csvHeaders[displayNameIndex]) < 0) {
        matchedColumn = csvHeaders[displayNameIndex];
      }

      // If not found, try internal name
      if (!matchedColumn) {
        const internalNameLower: string = field.internalName.toLowerCase();
        const internalNameIndex: number = lowerHeaders.indexOf(internalNameLower);
        if (internalNameIndex >= 0 && mappedHeaders.indexOf(csvHeaders[internalNameIndex]) < 0) {
          matchedColumn = csvHeaders[internalNameIndex];
        }
      }

      if (matchedColumn) {
        mappedHeaders.push(matchedColumn);
      }

      // Pre-populate default value from list field definition
      let defaultVal: string = field.defaultValue || '';
      // Translate SharePoint default value tokens to readable form
      if (field.fieldType === 'DateTime' && defaultVal === '[today]') {
        defaultVal = '[Heute]';
      }

      return {
        listField: field,
        csvColumn: matchedColumn,
        isKeyColumn: false,
        allowFillIn: false,
        defaultValue: defaultVal
      };
    });

    console.log(LOG, '_createInitialMappings \u2014 CSV headers:', csvHeaders,
      'mappings:', mappings.map((m: IFieldMapping) => ({
        field: m.listField.internalName,
        type: m.listField.fieldType,
        csvColumn: m.csvColumn || '(unmapped)'
      }))
    );

    return mappings;
  }

  // ─── Validation ──────────────────────────────────────────────────────

  private _allRequiredFieldsMapped(): boolean {
    return this.state.mappings
      .filter((m: IFieldMapping) => m.listField.required)
      .every((m: IFieldMapping) => {
        return m.csvColumn !== undefined ||
          (m.defaultValue !== undefined && m.defaultValue !== '');
      });
  }

  private _canImport(): boolean {
    return this.state.csvData !== undefined &&
      this.state.csvData.rows.length > 0 &&
      this.state.mappings.length > 0 &&
      this._allRequiredFieldsMapped();
  }

  // ─── Import Logic ────────────────────────────────────────────────────

  private _importData(): void {
    const { csvData, mappings, selectedWeb, selectedList } = this.state;
    if (!csvData || !selectedWeb || !selectedList) {
      return;
    }

    const webUrl: string = selectedWeb.url;
    const listId: string = selectedList.id;
    const activeMappings: IFieldMapping[] = mappings.filter(
      (m: IFieldMapping) => m.csvColumn !== undefined || (m.defaultValue !== undefined && m.defaultValue !== '')
    );
    const keyMapping: IFieldMapping | undefined = activeMappings.filter((m: IFieldMapping) => m.isKeyColumn)[0];

    // Taxonomy mappings that need special handling
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

    console.log(LOG, '_importData — activeMappings:', activeMappings.length,
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
    this.setState({ progress: progress });

    // Step 1: If we have a key column, look up existing items first
    const existingItemsPromise: Promise<{ [key: string]: number }> = keyMapping
      ? this._getExistingItemsMap(webUrl, listId, keyMapping, csvData)
      : Promise.resolve({});

    existingItemsPromise.then((existingItems: { [key: string]: number }) => {
      // Step 2: Process rows sequentially
      return this._processRows(
        csvData.rows,
        0,
        webUrl,
        listId,
        nonTaxonomyMappings,
        taxonomyMappings,
        csvData.headers,
        keyMapping,
        existingItems,
        progress
      );
    }).then(() => {
      progress.status = 'completed';
      this.setState({ progress: { ...progress } });
    // tslint:disable-next-line:no-any
    }).catch((error: Error) => {
      progress.status = 'error';
      progress.errorMessages.push(strings.ImportAborted + (error.message || error));
      this.setState({ progress: { ...progress } });
    });
  }

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

    return this.props.service.getExistingItems(webUrl, listId, keyMapping.listField.internalName, keyValues);
  }

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
    progress: IImportProgress
  ): Promise<void> {
    if (index >= rows.length) {
      return Promise.resolve();
    }

    const row: string[] = rows[index];

    return this._processRow(
      row, webUrl, listId, nonTaxMappings, taxMappings,
      headers, keyMapping, existingItems, progress
    ).then(() => {
      progress.current = index + 1;
      this.setState({ progress: { ...progress } });
      return this._processRows(
        rows, index + 1, webUrl, listId, nonTaxMappings,
        taxMappings, headers, keyMapping, existingItems, progress
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
    progress: IImportProgress
  ): Promise<void> {
    const service: CsvUploadService = this.props.service;
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
        csvValue = this._resolveDefaultValue(
          mapping.defaultValue, mapping.listField.fieldType
        );
      }

      return this._convertFieldWithRetry(
        webUrl, mapping, csvValue, rowNum
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
          console.log(LOG, '_processRow — processing taxonomy fields for itemId:', item.Id,
            'taxMappings:', taxMappings.length);
          return this._processTaxonomyFields(
            row, webUrl, listId, item.Id, taxMappings, headers, rowNum
          );
        } else {
          console.log(LOG, '_processRow — no taxonomy processing.',
            'taxMappings.length:', taxMappings.length,
            'item:', item ? 'Id=' + item.Id : '(undefined)');
        }
      });
    // tslint:disable-next-line:no-any
    }).catch((error: Error) => {
      progress.errors++;
      const friendlyMsg: string = this._extractErrorMessage(error);
      progress.errorMessages.push(
        strings.ErrorRowPrefix.replace('{0}', String(rowNum)) + friendlyMsg
      );
      if (progress.errorMessages.length > 100) {
        progress.errorMessages = progress.errorMessages.slice(0, 100);
        progress.errorMessages.push(strings.ErrorTruncated);
      }
    });
  }

  /**
   * Attempt to convert a field value; if conversion fails, show the
   * error dialog and let the user correct, skip the field, or skip the row.
   */
  // tslint:disable-next-line:no-any
  private _convertFieldWithRetry(
    webUrl: string,
    mapping: IFieldMapping,
    csvValue: string,
    rowNum: number
    // tslint:disable-next-line:no-any
  ): Promise<{ fieldName: string; value: any; skipRow?: boolean } | undefined> {
    const service: CsvUploadService = this.props.service;

    return service.convertFieldValue(
      webUrl, mapping.listField, csvValue, mapping.allowFillIn
    // tslint:disable-next-line:no-any
    ).catch((error: any) => {
      const friendlyMsg: string = this._extractErrorMessage(error);
      const errorInfo: IFieldErrorInfo = {
        rowNumber: rowNum,
        fieldDisplayName: mapping.listField.displayName || mapping.listField.internalName,
        fieldInternalName: mapping.listField.internalName,
        fieldType: mapping.listField.fieldType,
        csvValue: csvValue,
        errorMessage: friendlyMsg
      };

      return this._showFieldErrorDialog(errorInfo).then(
        (decision: FieldErrorDecision) => {
          if (decision.action === 'skip-row') {
            return { fieldName: '', value: undefined, skipRow: true };
          }
          if (decision.action === 'skip-field') {
            return undefined;
          }
          // 'use-value' — retry with the corrected value
          return this._convertFieldWithRetry(
            webUrl, mapping, decision.newValue, rowNum
          );
        }
      );
    });
  }

  private _processTaxonomyFields(
    row: string[],
    webUrl: string,
    listId: string,
    itemId: number,
    taxMappings: IFieldMapping[],
    headers: string[],
    rowNum: number
  ): Promise<void> {
    type ITaxonomyResult = { wssId: number; label: string; termGuid: string };
    const service: CsvUploadService = this.props.service;

    const processMapping: (idx: number) => Promise<void> = (idx: number) => {
      if (idx >= taxMappings.length) {
        return Promise.resolve();
      }

      const mapping: IFieldMapping = taxMappings[idx];
      const colIndex: number = headers.indexOf(mapping.csvColumn);
      let csvValue: string = colIndex >= 0 && colIndex < row.length ? row[colIndex] : '';
      // If cell is empty, fall back to the configured default value
      if ((!csvValue || csvValue === '') && mapping.defaultValue) {
        csvValue = this._resolveDefaultValue(mapping.defaultValue, mapping.listField.fieldType);
      }

      console.log(LOG, '_processTaxonomyFields — idx:', idx,
        'field:', mapping.listField.internalName,
        'type:', mapping.listField.fieldType,
        'csvColumn:', mapping.csvColumn,
        'colIndex:', colIndex,
        'csvValue:', JSON.stringify(csvValue),
        'termSetId:', mapping.listField.termSetId);

      if (!csvValue) {
        console.log(LOG, '_processTaxonomyFields — skipping (empty value)');
        return processMapping(idx + 1);
      }

      if (mapping.listField.fieldType === 'TaxonomyFieldTypeMulti') {
        const labels: string[] = this._parseTaxonomyExportValue(csvValue);
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
            console.log(LOG, '_processTaxonomyFields multi — labels:', labels,
              'resolved:', validResults.length, 'of', labels.length, validResults,
              'unresolved:', unresolvedLabels);

            // Resolve unresolved labels from the term store
            const storePromises: Promise<ITaxonomyResult | undefined>[] =
              unresolvedLabels.map((ul: string) =>
                service.resolveTermFromStore(webUrl, tsId, ul, mapping.listField.sspId)
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
                    '_processTaxonomyFields multi — terms not found:',
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
          return this._showFieldErrorDialog(errorInfo).then(
            (decision: FieldErrorDecision) => {
              if (decision.action === 'skip-row') {
                throw new Error('__SKIP_ROW__');
              }
              // skip-field or use-value: just continue to next mapping
              // (taxonomy re-resolve with a corrected value is complex;
              // the user can re-import to fix remaining issues)
            }
          );
        })
        .then(() => processMapping(idx + 1));
      }

      // Single taxonomy field — parse export format and take first label
      const parsedLabels: string[] = this._parseTaxonomyExportValue(csvValue);
      if (parsedLabels.length > 1) {
        console.warn(LOG, '_processTaxonomyFields single — multiple values found,',
          'using first label only:', JSON.stringify(parsedLabels[0]),
          'from:', JSON.stringify(csvValue));
      }
      csvValue = parsedLabels.length > 0 ? parsedLabels[0] : csvValue;
      const termSetId: string = mapping.listField.termSetId || '';
      const taxHiddenField: string = mapping.listField.taxonomyHiddenFieldName
        || mapping.listField.internalName;
      return service.resolveTaxonomyValue(webUrl, termSetId, csvValue)
        .then((result: ITaxonomyResult | undefined) => {
          console.log(LOG, '_processTaxonomyFields single — csvValue:', JSON.stringify(csvValue),
            'resolved:', result ? JSON.stringify(result) : '(not found, will search term store)');
          if (result) {
            return service.setTaxonomyFieldValue(
              webUrl, listId, itemId,
              mapping.listField.internalName, taxHiddenField,
              result.label, result.termGuid
            );
          }
          // Term not in TaxonomyHiddenList — search term store for the GUID
          return service.resolveTermFromStore(webUrl, termSetId, csvValue, mapping.listField.sspId)
            .then((storeResult: ITaxonomyResult | undefined) => {
              if (storeResult) {
                console.log(LOG, '_processTaxonomyFields single — resolved from STORE:',
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
              console.warn(LOG, '_processTaxonomyFields — term not found anywhere:',
                JSON.stringify(csvValue),
                'termSetId:', termSetId,
                'sspId:', mapping.listField.sspId);
              throw new Error(
                strings.ErrorTaxonomyTermNotFound
                  .replace('{0}', csvValue)
                  .replace('{1}', mapping.listField.displayName || mapping.listField.internalName)
              );
            });
        })
        // tslint:disable-next-line:no-any
        .catch((err: any) => {
          // Translate SharePoint SPFieldValueException into a friendly message
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
          return this._showFieldErrorDialog(errorInfo).then(
            (decision: FieldErrorDecision) => {
              if (decision.action === 'skip-row') {
                throw new Error('__SKIP_ROW__');
              }
              // skip-field or use-value: continue to next mapping
            }
          );
        })
        .then(() => processMapping(idx + 1));
    };

    return processMapping(0);
  }

  // ─── Helpers ─────────────────────────────────────────────────────────

  /**
   * Parse a taxonomy value that may be in SharePoint export format.
   *
   * SharePoint's "Export to Excel" writes taxonomy fields as:
   *   Single:  "<wssId>;#<label>"          e.g. "6;#Einsatz"
   *   Multi:   "<wssId>;#<label>;#<wssId>;#<label>"  e.g. "12;#Lehre;#13;#Forschung"
   *
   * SharePoint sometimes prefixes labels with an extra "#" in the
   * export (e.g. "22;##IT-Verfahren" instead of "22;#IT-Verfahren").
   * After splitting on ";#", this leaves a leading "#" on the label
   * which must be stripped.
   *
   * This method detects the pattern and returns an array of clean labels.
   * If the value is NOT in export format, it splits by ";" as usual.
   */
  private _parseTaxonomyExportValue(csvValue: string): string[] {
    // Detect SharePoint export format: starts with digits followed by ;#
    if (/^\d+;#/.test(csvValue)) {
      const segments: string[] = csvValue.split(';#');
      const labels: string[] = [];
      for (let i: number = 0; i < segments.length; i++) {
        let seg: string = segments[i].trim();
        if (!seg) {
          continue;
        }
        // Skip segments that are purely numeric (wssId values)
        if (/^\d+$/.test(seg)) {
          continue;
        }
        // Strip leading "#" left over from SharePoint's "##" export artifact
        if (seg.charAt(0) === '#') {
          seg = seg.substring(1);
        }
        if (seg) {
          labels.push(seg);
        }
      }
      console.log(LOG, '_parseTaxonomyExportValue — detected SP export format.',
        'raw:', JSON.stringify(csvValue), 'labels:', labels);
      return labels;
    }
    // Plain format: split by semicolon
    return csvValue.split(';').map((s: string) => s.trim()).filter((s: string) => s !== '');
  }

  /**
   * Resolve special tokens in default values (e.g. [Heute] / [Today] for date fields).
   */
  private _resolveDefaultValue(defaultValue: string, fieldType: string): string {
    if (!defaultValue) {
      return '';
    }
    const lower: string = defaultValue.toLowerCase().trim();

    // Date tokens
    if (fieldType === 'DateTime') {
      if (lower === '[heute]' || lower === '[today]') {
        return new Date().toISOString();
      }
      // Handle [Heute]+N / [Today]+N  (offset in days)
      const offsetMatch: RegExpMatchArray = lower.match(/^\[(heute|today)\]\s*([+-]\s*\d+)$/);
      if (offsetMatch) {
        const offsetDays: number = parseInt(offsetMatch[2].replace(/\s/g, ''), 10);
        const d: Date = new Date();
        d.setDate(d.getDate() + offsetDays);
        return d.toISOString();
      }
    }

    // Boolean tokens
    if (fieldType === 'Boolean') {
      if (lower === '1' || lower === 'true' || lower === 'ja' || lower === 'yes') {
        return 'true';
      }
      if (lower === '0' || lower === 'false' || lower === 'nein' || lower === 'no') {
        return 'false';
      }
    }

    return defaultValue;
  }

  /**
   * Extract a user-friendly error message from a SharePoint REST error
   * or any other Error object. Tries to pull out the odata.error message
   * value from the JSON body; falls back to error.message.
   */
  private _extractErrorMessage(error: Error): string {
    const raw: string = error && error.message ? error.message : String(error);
    // Try to extract the odata.error JSON from PnPJS error messages
    // Format: "Error making HttpClient request ... ::> { JSON }"
    const jsonStart: number = raw.indexOf('::>');
    if (jsonStart >= 0) {
      const jsonPart: string = raw.substring(jsonStart + 3).trim();
      try {
        const parsed: { 'odata.error'?: { message?: { value?: string } } } =
          JSON.parse(jsonPart);
        if (parsed['odata.error'] &&
            parsed['odata.error'].message &&
            parsed['odata.error'].message.value) {
          return parsed['odata.error'].message.value;
        }
      } catch (e) {
        // JSON parse failed — fall through to raw message
      }
    }
    return raw;
  }

  private _resetProgress(): IImportProgress {
    return {
      total: 0,
      current: 0,
      created: 0,
      updated: 0,
      errors: 0,
      errorMessages: [],
      status: 'idle'
    };
  }
}
