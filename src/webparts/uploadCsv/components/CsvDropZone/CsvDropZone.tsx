import * as React from 'react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import styles from '../UploadCsv.module.scss';
import * as strings from 'UploadCsvWebPartStrings';

export interface ICsvDropZoneProps {
  onFileSelected: (file: File) => void;
  fileName: string | undefined;
}

export interface ICsvDropZoneState {
  isDragOver: boolean;
}

export default class CsvDropZone extends React.Component<ICsvDropZoneProps, ICsvDropZoneState> {
  private _fileInput: HTMLInputElement;

  constructor(props: ICsvDropZoneProps) {
    super(props);
    this.state = {
      isDragOver: false
    };
  }

  public render(): React.ReactElement<ICsvDropZoneProps> {
    const dropZoneClass: string = this.state.isDragOver
      ? `${styles.dropZone} ${styles.dropZoneActive}`
      : styles.dropZone;

    return (
      <div
        className={dropZoneClass}
        onDragOver={this._onDragOver}
        onDragLeave={this._onDragLeave}
        onDrop={this._onDrop}
        onClick={this._onClick}
        role='button'
        aria-expanded={false}
        aria-controls='csvFileInput'
      >
        <input
          id='csvFileInput'
          type='file'
          accept='.csv,.txt'
          ref={this._setFileInputRef}
          style={{ display: 'none' }}
          onChange={this._onFileInputChange}
        />
        <Icon iconName='BulkUpload' className={styles.dropZoneIcon} />
        {this.props.fileName ? (
          <div className={styles.dropZoneText}>
            <strong>{this.props.fileName}</strong>
            <br />
            {strings.DropZoneHintReplace}
          </div>
        ) : (
          <div className={styles.dropZoneText}>
            {strings.DropZoneHint}
          </div>
        )}
      </div>
    );
  }

  private _setFileInputRef = (ref: HTMLInputElement): void => {
    this._fileInput = ref;
  }

  private _onDragOver = (event: React.DragEvent<HTMLDivElement>): void => {
    event.preventDefault();
    event.stopPropagation();
    this.setState({ isDragOver: true });
  }

  private _onDragLeave = (event: React.DragEvent<HTMLDivElement>): void => {
    event.preventDefault();
    event.stopPropagation();
    this.setState({ isDragOver: false });
  }

  private _onDrop = (event: React.DragEvent<HTMLDivElement>): void => {
    event.preventDefault();
    event.stopPropagation();
    this.setState({ isDragOver: false });

    if (event.dataTransfer.files && event.dataTransfer.files.length > 0) {
      const file: File = event.dataTransfer.files[0];
      if (file.name && file.name.length > 4) {
        const lowerName: string = file.name.toLowerCase();
        const isCsv: boolean = lowerName.indexOf('.csv') === lowerName.length - 4;
        const isTxt: boolean = lowerName.indexOf('.txt') === lowerName.length - 4;
        if (isCsv || isTxt) {
          this.props.onFileSelected(file);
        }
      }
    }
  }

  private _onClick = (): void => {
    if (this._fileInput) {
      // Reset value so re-selecting the same file triggers onChange
      this._fileInput.value = '';
      this._fileInput.click();
    }
  }

  private _onFileInputChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    if (event.target.files && event.target.files.length > 0) {
      this.props.onFileSelected(event.target.files[0]);
    }
  }
}
