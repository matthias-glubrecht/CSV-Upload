import { IListField, IFieldMapping } from '../models';

const LOG: string = '[CsvUpload]';

/**
 * Auto-match CSV headers to SharePoint list fields by display name
 * or internal name (case-insensitive). Each CSV header is used at
 * most once.
 */
export function createInitialMappings(
  fields: IListField[], csvHeaders: string[]
): IFieldMapping[] {
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

  console.log(LOG, 'createInitialMappings \u2014 CSV headers:', csvHeaders,
    'mappings:', mappings.map((m: IFieldMapping) => ({
      field: m.listField.internalName,
      type: m.listField.fieldType,
      csvColumn: m.csvColumn || '(unmapped)'
    }))
  );

  return mappings;
}

/**
 * Check whether all required SharePoint fields have either a
 * mapped CSV column or a non-empty default value.
 */
export function allRequiredFieldsMapped(mappings: IFieldMapping[]): boolean {
  return mappings
    .filter((m: IFieldMapping) => m.listField.required)
    .every((m: IFieldMapping) => {
      return m.csvColumn !== undefined ||
        (m.defaultValue !== undefined && m.defaultValue !== '');
    });
}
