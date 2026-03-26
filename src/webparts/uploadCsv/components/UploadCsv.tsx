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
  IImportProgress
} from '../models';
import CsvUploadService from '../service/CsvUploadService';
import { parseCsv, detectDelimiter } from '../utils/CsvParser';
import SiteCollectionPicker from './SiteCollectionPicker/SiteCollectionPicker';
import WebPicker from './WebPicker/WebPicker';
import ListPicker from './ListPicker/ListPicker';
import CsvDropZone from './CsvDropZone/CsvDropZone';
import MappingTable from './MappingTable/MappingTable';
import ImportProgress from './ImportProgress/ImportProgress';
import * as strings from 'UploadCsvWebPartStrings';

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
}

export default class UploadCsv extends React.Component<IUploadCsvProps, IUploadCsvState> {

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
      errorMessage: undefined
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
      const csvText: string = (e.target as any).result as string;
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
    reader.readAsText(file, 'UTF-8');
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
      errorMessage: undefined
    });
  }

  private _clearError = (): void => {
    this.setState({ errorMessage: undefined });
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

    // tslint:disable-next-line:no-any
    type FieldResult = { fieldName: string; value: any };
    // Build field values from non-taxonomy mappings
    const fieldValuePromises: Promise<FieldResult | undefined>[] =
      nonTaxMappings.map((mapping: IFieldMapping) => {
        const colIndex: number = headers.indexOf(mapping.csvColumn);
        let csvValue: string = colIndex >= 0 && colIndex < row.length
          ? row[colIndex] : '';
        // If cell is empty, fall back to the configured default value
        if ((!csvValue || csvValue === '') && mapping.defaultValue) {
          csvValue = this._resolveDefaultValue(
            mapping.defaultValue, mapping.listField.fieldType
          );
        }
        return service.convertFieldValue(
          webUrl, mapping.listField, csvValue, mapping.allowFillIn
        );
      });

    return Promise.all(fieldValuePromises)
      .then((fieldResults: (FieldResult | undefined)[]) => {
      // tslint:disable-next-line:no-any
      const fieldValues: { [key: string]: any } = {};
      fieldResults.forEach((result: FieldResult | undefined) => {
        if (result) {
          fieldValues[result.fieldName] = result.value;
        }
      });

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
          return this._processTaxonomyFields(
            row, webUrl, listId, item.Id, taxMappings, headers
          );
        }
      });
    // tslint:disable-next-line:no-any
    }).catch((error: Error) => {
      progress.errors++;
      const rowNum: number = progress.current + 1;
      progress.errorMessages.push(strings.ErrorRowPrefix.replace('{0}', String(rowNum)) + (error.message || error));
      // Limit error messages to prevent memory issues
      if (progress.errorMessages.length > 100) {
        progress.errorMessages = progress.errorMessages.slice(0, 100);
        progress.errorMessages.push(strings.ErrorTruncated);
      }
    });
  }

  private _processTaxonomyFields(
    row: string[],
    webUrl: string,
    listId: string,
    itemId: number,
    taxMappings: IFieldMapping[],
    headers: string[]
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

      if (!csvValue) {
        return processMapping(idx + 1);
      }

      if (mapping.listField.fieldType === 'TaxonomyFieldTypeMulti') {
        const labels: string[] = csvValue.split(';').map((s: string) => s.trim());
        const tsId: string = mapping.listField.termSetId || '';
        const termPromises: Promise<ITaxonomyResult | undefined>[] =
          labels.map((label: string) =>
            service.resolveTaxonomyValue(webUrl, tsId, label)
          );
        return Promise.all(termPromises).then(
          (results: (ITaxonomyResult | undefined)[]) => {
            const validResults: ITaxonomyResult[] = results.filter(
              (r: ITaxonomyResult | undefined) => r !== undefined
            ) as ITaxonomyResult[];
            if (validResults.length > 0) {
              return service.setTaxonomyMultiFieldValue(
                webUrl, listId, itemId,
                mapping.listField.internalName, validResults
              );
            }
          }
        ).then(() => processMapping(idx + 1));
      }

      // Single taxonomy field
      const termSetId: string = mapping.listField.termSetId || '';
      return service.resolveTaxonomyValue(webUrl, termSetId, csvValue)
        .then((result: ITaxonomyResult | undefined) => {
          if (result) {
            return service.setTaxonomyFieldValue(
              webUrl, listId, itemId,
              mapping.listField.internalName,
              result.wssId, result.label, result.termGuid
            );
          }
        }).then(() => processMapping(idx + 1));
    };

    return processMapping(0);
  }

  // ─── Helpers ─────────────────────────────────────────────────────────

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
