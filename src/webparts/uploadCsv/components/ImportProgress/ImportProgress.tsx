import * as React from 'react';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import styles from '../UploadCsv.module.scss';
import { IImportProgress } from '../../models';
import { formatString } from '../../utils/ImportHelpers';
import * as strings from 'UploadCsvWebPartStrings';

export interface IImportProgressProps {
  progress: IImportProgress;
}

export default class ImportProgress extends React.Component<IImportProgressProps, {}> {

  public render(): React.ReactElement<IImportProgressProps> {
    const { progress } = this.props;

    if (progress.status === 'idle') {
      return <div />;
    }

    const percentComplete: number = progress.total > 0 ? progress.current / progress.total : 0;
    const description: string = formatString(
      strings.ProgressDescription,
      progress.current, progress.total, progress.created, progress.updated, progress.errors
    );

    return (
      <div className={styles.progressContainer}>
        <ProgressIndicator
          label={progress.status === 'completed' ? strings.ImportCompletedLabel : strings.ImportRunningLabel}
          description={description}
          percentComplete={percentComplete}
        />

        {progress.status === 'completed' && progress.errors === 0 && (
          <MessageBar messageBarType={MessageBarType.success} className={styles.messageBar}>
            {formatString(strings.ImportSuccessMessage, progress.total, progress.created, progress.updated)}
          </MessageBar>
        )}

        {progress.status === 'completed' && progress.errors > 0 && (
          <MessageBar messageBarType={MessageBarType.warning} className={styles.messageBar}>
            {formatString(strings.ImportWarningMessage, progress.errors, progress.created, progress.updated)}
          </MessageBar>
        )}

        {progress.status === 'error' && (
          <MessageBar messageBarType={MessageBarType.error} className={styles.messageBar}>
            {strings.ImportErrorMessage}
          </MessageBar>
        )}

        {progress.errorMessages.length > 0 && (
          <div className={styles.errorList}>
            <strong>{strings.ErrorMessagesTitle}</strong>
            <ul>
              {progress.errorMessages.map((msg: string, idx: number) => (
                <li key={idx}>{msg}</li>
              ))}
            </ul>
          </div>
        )}
      </div>
    );
  }
}
