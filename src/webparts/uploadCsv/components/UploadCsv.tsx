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
import ImportEngine from '../service/ImportEngine';
import { parseCsv, detectDelimiter, decodeCsvBytes } from '../utils/CsvParser';
import { resetProgress } from '../utils/ImportHelpers';
import { createInitialMappings, allRequiredFieldsMapped } from '../utils/MappingHelpers';
import SiteCollectionPicker from './SiteCollectionPicker/SiteCollectionPicker';
import WebPicker from './WebPicker/WebPicker';
import ListPicker from './ListPicker/ListPicker';
import CsvDropZone from './CsvDropZone/CsvDropZone';
import MappingTable from './MappingTable/MappingTable';
import ImportProgress from './ImportProgress/ImportProgress';
import FieldErrorDialog from './FieldErrorDialog/FieldErrorDialog';
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
  /** When set, the FieldErrorDialog is shown and the import is paused */
  fieldError: IFieldErrorInfo | undefined;
}

export default class UploadCsv extends React.Component<IUploadCsvProps, IUploadCsvState> {

  /** Stored resolve function for the field error dialog promise */
  private _fieldErrorResolve: ((decision: FieldErrorDecision) => void) | undefined;

  /** Import engine instance, created once in constructor */
  private _importEngine: ImportEngine;

  constructor(props: IUploadCsvProps) {
    super(props);

    this._importEngine = new ImportEngine(props.service);

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
      progress: resetProgress(),
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
              {!allRequiredFieldsMapped(mappings) && (
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
      progress: resetProgress()
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
      progress: resetProgress()
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
      progress: resetProgress()
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

      const mappings: IFieldMapping[] = createInitialMappings(this.state.listFields, csvData.headers);

      this.setState({
        csvData: csvData,
        csvFileName: file.name,
        mappings: mappings,
        progress: resetProgress()
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
    const { csvData, mappings, selectedWeb, selectedList } = this.state;
    if (!csvData || !selectedWeb || !selectedList) {
      return;
    }

    this._importEngine.run(
      csvData,
      mappings,
      selectedWeb.url,
      selectedList.id,
      {
        onProgress: (progress: IImportProgress) => {
          this.setState({ progress: progress });
        },
        onFieldError: (error: IFieldErrorInfo) => {
          return this._showFieldErrorDialog(error);
        }
      }
    );
  }

  private _onReset = (): void => {
    this.setState({
      csvData: undefined,
      csvFileName: undefined,
      mappings: [],
      progress: resetProgress(),
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

  // ─── Validation ──────────────────────────────────────────────────────

  private _canImport(): boolean {
    return this.state.csvData !== undefined &&
      this.state.csvData.rows.length > 0 &&
      this.state.mappings.length > 0 &&
      allRequiredFieldsMapped(this.state.mappings);
  }
}
