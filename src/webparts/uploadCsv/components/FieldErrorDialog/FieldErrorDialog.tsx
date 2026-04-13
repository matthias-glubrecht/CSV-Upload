import * as React from 'react';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { IFieldErrorInfo, FieldErrorDecision } from '../../models';
import styles from '../UploadCsv.module.scss';
import * as strings from 'UploadCsvWebPartStrings';

export interface IFieldErrorDialogProps {
  /** The error information to display */
  error: IFieldErrorInfo;
  /** Callback when the user makes a decision */
  onDecision: (decision: FieldErrorDecision) => void;
}

export interface IFieldErrorDialogState {
  correctedValue: string;
}

export default class FieldErrorDialog
  extends React.Component<IFieldErrorDialogProps, IFieldErrorDialogState> {

  constructor(props: IFieldErrorDialogProps) {
    super(props);
    this.state = {
      correctedValue: props.error.csvValue
    };
  }

  public componentWillReceiveProps(nextProps: IFieldErrorDialogProps): void {
    if (nextProps.error !== this.props.error) {
      this.setState({ correctedValue: nextProps.error.csvValue });
    }
  }

  public render(): React.ReactElement<IFieldErrorDialogProps> {
    const { error } = this.props;
    const { correctedValue } = this.state;

    return (
      <Dialog
        hidden={false}
        onDismiss={this._onSkipRow}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: strings.FieldErrorDialogTitle
        }}
        modalProps={{
          isBlocking: true,
          containerClassName: styles.fieldErrorDialogContainer
        }}
      >
        <div className={styles.fieldErrorDialogContent}>
          <div className={styles.fieldErrorDialogInfoGrid}>
            <Label>{strings.FieldErrorDialogRowLabel}:</Label>
            <span>{error.rowNumber}</span>

            <Label>{strings.FieldErrorDialogFieldLabel}:</Label>
            <span>{error.fieldDisplayName}</span>

            <Label>{strings.FieldErrorDialogValueLabel}:</Label>
            <span className={styles.fieldErrorDialogOriginalValue}>{error.csvValue}</span>

            <Label>{strings.FieldErrorDialogErrorLabel}:</Label>
            <span className={styles.fieldErrorDialogErrorText}>{error.errorMessage}</span>
          </div>

          <TextField
            label={strings.FieldErrorDialogCorrectedValueLabel}
            value={correctedValue}
            onChanged={this._onCorrectedValueChanged}
            placeholder={strings.FieldErrorDialogCorrectedValuePlaceholder}
          />

          {/* Action buttons — three equal-width in a row */}
          <div className={styles.fieldErrorDialogActions}>
            <PrimaryButton
              text={strings.FieldErrorDialogUseValueButton}
              onClick={this._onUseValue}
              iconProps={{ iconName: 'CheckMark' }}
              className={styles.fieldErrorDialogActionButton}
            />
            <DefaultButton
              text={strings.FieldErrorDialogSkipFieldButton}
              onClick={this._onSkipField}
              iconProps={{ iconName: 'Forward' }}
              className={styles.fieldErrorDialogActionButton}
            />
            <DefaultButton
              text={strings.FieldErrorDialogSkipRowButton}
              onClick={this._onSkipRow}
              iconProps={{ iconName: 'Cancel' }}
              className={styles.fieldErrorDialogActionButton}
            />
          </div>

          {/* Abort — visually separated */}
          <div className={styles.fieldErrorDialogAbortRow}>
            <DefaultButton
              text={strings.FieldErrorDialogAbortImportButton}
              onClick={this._onAbortImport}
              iconProps={{ iconName: 'StopSolid' }}
              className={styles.fieldErrorDialogAbortButton}
            />
          </div>
        </div>

        <DialogFooter>
          {/* Empty — buttons are rendered above for custom layout */}
        </DialogFooter>
      </Dialog>
    );
  }

  private _onCorrectedValueChanged = (newValue: string): void => {
    this.setState({ correctedValue: newValue });
  }

  private _onUseValue = (): void => {
    this.props.onDecision({
      action: 'use-value',
      newValue: this.state.correctedValue
    });
  }

  private _onSkipField = (): void => {
    this.props.onDecision({ action: 'skip-field' });
  }

  private _onSkipRow = (): void => {
    this.props.onDecision({ action: 'skip-row' });
  }

  private _onAbortImport = (): void => {
    this.props.onDecision({ action: 'abort-import' });
  }
}
