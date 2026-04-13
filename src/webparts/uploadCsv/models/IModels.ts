/**
 * Represents a SharePoint site collection.
 */
export interface ISiteCollection {
  title: string;
  url: string;
}

/**
 * Represents a SharePoint web (sub-site).
 */
export interface IWeb {
  title: string;
  url: string;
  id: string;
}

/**
 * Represents a SharePoint list (non-library).
 */
export interface IListInfo {
  title: string;
  id: string;
}

/**
 * Supported SharePoint field types for mapping.
 */
export type SpFieldType =
  | 'Text'
  | 'Note'
  | 'Number'
  | 'DateTime'
  | 'Choice'
  | 'MultiChoice'
  | 'Lookup'
  | 'LookupMulti'
  | 'User'
  | 'UserMulti'
  | 'Boolean'
  | 'Currency'
  | 'URL'
  | 'TaxonomyFieldType'
  | 'TaxonomyFieldTypeMulti'
  | 'Unknown';

/**
 * Represents a SharePoint list field.
 */
export interface IListField {
  internalName: string;
  displayName: string;
  fieldType: SpFieldType;
  required: boolean;
  /** For choice fields: available choices */
  choices?: string[];
  /** For lookup fields: lookup list ID */
  lookupListId?: string;
  /** For lookup fields: lookup field name */
  lookupFieldName?: string;
  /** For lookup fields: lookup web id */
  lookupWebId?: string;
  /** For taxonomy fields: term set id */
  termSetId?: string;
  /** For taxonomy fields: term store id (SspId) */
  sspId?: string;
  /** For taxonomy fields: the hidden note field internal name (e.g. TaxCatchAll) */
  taxonomyHiddenFieldName?: string;
  /** Default value from the list field definition */
  defaultValue?: string;
}

/**
 * A mapping from a SharePoint field to a CSV column.
 */
export interface IFieldMapping {
  /** The SharePoint list field */
  listField: IListField;
  /** The CSV column header that maps to this field (undefined = unmapped) */
  csvColumn: string | undefined;
  /** Whether this mapping is the key column for upsert */
  isKeyColumn: boolean;
  /** For choice fields: allow values not in the configured choices */
  allowFillIn: boolean;
  /** Default value to use when CSV cell is empty (editable, pre-populated from list definition) */
  defaultValue: string;
}

/**
 * Parsed CSV data.
 */
export interface ICsvData {
  headers: string[];
  rows: string[][];
}

/**
 * Progress information during import.
 */
export interface IImportProgress {
  total: number;
  current: number;
  created: number;
  updated: number;
  errors: number;
  errorMessages: string[];
  status: 'idle' | 'running' | 'completed' | 'error';
}

/**
 * Describes a field-level error during import that the user can correct.
 */
export interface IFieldErrorInfo {
  /** Row number (1-based) where the error occurred */
  rowNumber: number;
  /** Display name of the SharePoint field */
  fieldDisplayName: string;
  /** Internal name of the SharePoint field */
  fieldInternalName: string;
  /** Field type (Choice, Lookup, User, etc.) */
  fieldType: SpFieldType;
  /** The CSV value that caused the error */
  csvValue: string;
  /** User-friendly error description */
  errorMessage: string;
}

/**
 * User decision for how to handle a field error.
 */
export type FieldErrorDecision =
  | { action: 'use-value'; newValue: string }
  | { action: 'skip-field' }
  | { action: 'skip-row' }
  | { action: 'abort-import' };
