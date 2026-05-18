// tslint:disable:export-name
// File name is intentionally lowercase to match the utils/ naming
// convention; the exported constant name follows the all-caps
// convention for module-level string constants.

/**
 * Shared console log prefix used by all CSV-Upload modules.
 * Imported by services and utilities that emit diagnostic
 * warnings or errors via console.warn / console.error.
 */
export const LOG: string = '[CsvUpload]';
