import * as React from 'react';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import styles from '../UploadCsv.module.scss';
import { IFieldMapping } from '../../models';
import * as strings from 'UploadCsvWebPartStrings';

export interface IMappingTableProps {
  mappings: IFieldMapping[];
  csvHeaders: string[];
  onMappingChanged: (updatedMappings: IFieldMapping[]) => void;
}

export default class MappingTable extends React.Component<IMappingTableProps, {}> {

  public render(): React.ReactElement<IMappingTableProps> {
    const { mappings } = this.props;
    return (
      <div className={styles.mappingTableContainer}>
        <table className={styles.mappingTable}>
          <thead>
            <tr>
              <th className={styles.mappingTableHeader}>{strings.HeaderKeyColumn}</th>
              <th className={styles.mappingTableHeader}>{strings.HeaderSharePointField}</th>
              <th className={styles.mappingTableHeader}>{strings.HeaderFieldType}</th>
              <th className={styles.mappingTableHeader}>{strings.HeaderRequired}</th>
              <th className={styles.mappingTableHeader}>{strings.HeaderDefaultValue}</th>
              <th className={styles.mappingTableHeader}>{strings.HeaderCsvColumn}</th>
              <th className={styles.mappingTableHeader}>{strings.HeaderOptions}</th>
            </tr>
          </thead>
          <tbody>
            {mappings.map((mapping: IFieldMapping, index: number) => {
              return this._renderMappingRow(mapping, index);
            })}
          </tbody>
        </table>
      </div>
    );
  }

  private _renderMappingRow(mapping: IFieldMapping, index: number): React.ReactElement<{}> {
    const csvOptions: IDropdownOption[] = this._getCsvOptions(mapping);
    const fieldTypeLabel: string = this._getFieldTypeLabel(mapping.listField.fieldType);
    const isChoiceField: boolean =
      mapping.listField.fieldType === 'Choice' ||
      mapping.listField.fieldType === 'MultiChoice';

    // SharePoint cannot use these field types in OData filter
    // expressions, so they must not be used as key columns.
    const canBeKeyColumn: boolean =
      mapping.listField.fieldType !== 'TaxonomyFieldType' &&
      mapping.listField.fieldType !== 'TaxonomyFieldTypeMulti' &&
      mapping.listField.fieldType !== 'LookupMulti' &&
      mapping.listField.fieldType !== 'UserMulti' &&
      mapping.listField.fieldType !== 'MultiChoice' &&
      mapping.listField.fieldType !== 'Note';

    const rowClass: string = mapping.listField.required && !mapping.csvColumn
      ? styles.mappingRowRequired
      : styles.mappingRow;

    return (
      <tr key={mapping.listField.internalName} className={rowClass}>
        <td className={styles.mappingTableCell}>
          <Checkbox
            checked={mapping.isKeyColumn}
            disabled={!canBeKeyColumn}
            onChange={
              (ev: React.FormEvent<HTMLElement>, checked: boolean) =>
                this._onKeyColumnChanged(index, checked)
            }
            ariaLabel={strings.KeyColumnAriaLabel + ' ' + mapping.listField.displayName}
          />
        </td>
        <td className={styles.mappingTableCell}>
          <span className={styles.fieldName}>{mapping.listField.displayName}</span>
          <span className={styles.fieldInternalName}>({mapping.listField.internalName})</span>
        </td>
        <td className={styles.mappingTableCell}>
          <span className={styles.fieldType}>{fieldTypeLabel}</span>
        </td>
        <td className={styles.mappingTableCell}>
          <span className={mapping.listField.required ? styles.requiredYes : styles.requiredNo}>
            {mapping.listField.required ? strings.Yes : strings.No}
          </span>
        </td>
        <td className={styles.mappingTableCell}>
          {mapping.listField.required && !mapping.isKeyColumn ? (
            <TextField
              value={mapping.defaultValue}
              onChanged={(newValue: string) => this._onDefaultValueChanged(index, newValue)}
              placeholder={strings.DefaultValuePlaceholder}
            />
          ) : undefined}
        </td>
        <td className={styles.mappingTableCell}>
          <Dropdown
            placeHolder={strings.NotMapped}
            options={csvOptions}
            selectedKey={mapping.csvColumn || ''}
            onChanged={(option: IDropdownOption) => this._onCsvColumnChanged(index, option)}
          />
        </td>
        <td className={styles.mappingTableCell}>
          {isChoiceField && mapping.csvColumn ? (
            <Toggle
              label={strings.AllowFillInLabel}
              checked={mapping.allowFillIn}
              onChanged={(checked: boolean) => this._onAllowFillInChanged(index, checked)}
              onText={strings.Yes}
              offText={strings.No}
            />
          ) : undefined}
        </td>
      </tr>
    );
  }

  private _getCsvOptions(mapping: IFieldMapping): IDropdownOption[] {
    const { csvHeaders } = this.props;

    const options: IDropdownOption[] = [
      { key: '', text: strings.NotMapped }
    ];

    csvHeaders.forEach((header: string) => {
      options.push({
        key: header,
        text: header
      });
    });

    return options;
  }

  private _onCsvColumnChanged = (index: number, option: IDropdownOption): void => {
    const updatedMappings: IFieldMapping[] = this.props.mappings.map(
      (m: IFieldMapping, i: number) => {
        if (i === index) {
          return {
            ...m,
            csvColumn: option.key === '' ? undefined : option.key as string
          };
        }
        return m;
      }
    );
    this.props.onMappingChanged(updatedMappings);
  }

  private _onKeyColumnChanged = (index: number, checked: boolean): void => {
    const updatedMappings: IFieldMapping[] = this.props.mappings.map(
      (m: IFieldMapping, i: number) => {
        if (i === index) {
          return { ...m, isKeyColumn: checked };
        }
        // Only one key column at a time
        if (checked) {
          return { ...m, isKeyColumn: false };
        }
        return m;
      }
    );
    this.props.onMappingChanged(updatedMappings);
  }

  private _onDefaultValueChanged = (index: number, newValue: string): void => {
    const updatedMappings: IFieldMapping[] = this.props.mappings.map(
      (m: IFieldMapping, i: number) => {
        if (i === index) {
          return { ...m, defaultValue: newValue };
        }
        return m;
      }
    );
    this.props.onMappingChanged(updatedMappings);
  }

  private _onAllowFillInChanged = (index: number, checked: boolean): void => {
    const updatedMappings: IFieldMapping[] = this.props.mappings.map(
      (m: IFieldMapping, i: number) => {
        if (i === index) {
          return { ...m, allowFillIn: checked };
        }
        return m;
      }
    );
    this.props.onMappingChanged(updatedMappings);
  }

  private _getFieldTypeLabel(fieldType: string): string {
    const labels: { [key: string]: string } = {
      'Text': strings.FieldTypeText,
      'Note': strings.FieldTypeNote,
      'Number': strings.FieldTypeNumber,
      'Currency': strings.FieldTypeCurrency,
      'DateTime': strings.FieldTypeDateTime,
      'Choice': strings.FieldTypeChoice,
      'MultiChoice': strings.FieldTypeMultiChoice,
      'Lookup': strings.FieldTypeLookup,
      'LookupMulti': strings.FieldTypeLookupMulti,
      'User': strings.FieldTypeUser,
      'UserMulti': strings.FieldTypeUserMulti,
      'Boolean': strings.FieldTypeBoolean,
      'URL': strings.FieldTypeURL,
      'TaxonomyFieldType': strings.FieldTypeTaxonomy,
      'TaxonomyFieldTypeMulti': strings.FieldTypeTaxonomyMulti
    };
    return labels[fieldType] || fieldType;
  }
}
