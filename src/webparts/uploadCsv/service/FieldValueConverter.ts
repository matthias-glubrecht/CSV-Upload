// tslint:disable:no-any
import CsvUploadService from './CsvUploadService';
import { IListField } from '../models';
import { formatString } from '../utils';
import * as strings from 'UploadCsvWebPartStrings';

/** Shape returned by convertFieldValue for the SharePoint REST update payload. */
export interface IFieldValueResult {
  fieldName: string;
  value: any;
}

/**
 * Converts CSV string values into the payload shape expected by
 * the SharePoint REST API for each supported field type. Lookup
 * and user resolution is delegated back to CsvUploadService.
 *
 * Taxonomy fields are deliberately skipped here — they are
 * processed separately by TaxonomyProcessor after the item has
 * been created/updated.
 */
export default class FieldValueConverter {

  private _service: CsvUploadService;

  constructor(service: CsvUploadService) {
    this._service = service;
  }

  /**
   * Convert a CSV string value to the appropriate field value
   * for the SharePoint REST API. Returns `undefined` when the
   * value is empty or when the field type is handled elsewhere
   * (e.g. taxonomy fields).
   */
  public convertFieldValue(
    webUrl: string,
    field: IListField,
    csvValue: string,
    allowChoiceFillIn: boolean
  ): Promise<IFieldValueResult | undefined> {

    if (csvValue === undefined || csvValue === '') {
      return Promise.resolve(undefined);
    }

    switch (field.fieldType) {
      case 'Text':
      case 'Note':
        return Promise.resolve({
          fieldName: field.internalName, value: csvValue
        });

      case 'Number':
      case 'Currency': {
        const numVal: number = parseFloat(csvValue.replace(',', '.'));
        if (isNaN(numVal)) {
          return Promise.resolve(undefined);
        }
        return Promise.resolve({
          fieldName: field.internalName, value: numVal
        });
      }

      case 'Boolean':
        return Promise.resolve({
          fieldName: field.internalName,
          value: this._parseBooleanValue(csvValue)
        });

      case 'DateTime': {
        const dateVal: string | undefined = this._parseDateValue(csvValue);
        if (!dateVal) {
          return Promise.resolve(undefined);
        }
        return Promise.resolve({
          fieldName: field.internalName, value: dateVal
        });
      }

      case 'Choice':
        return this._convertChoice(field, csvValue, allowChoiceFillIn);

      case 'MultiChoice':
        return this._convertMultiChoice(field, csvValue, allowChoiceFillIn);

      case 'Lookup':
        return this._convertLookup(webUrl, field, csvValue);

      case 'LookupMulti':
        return this._convertLookupMulti(webUrl, field, csvValue);

      case 'User':
        return this._convertUser(webUrl, field, csvValue);

      case 'UserMulti':
        return this._convertUserMulti(webUrl, field, csvValue);

      case 'URL':
        return Promise.resolve(this._convertUrl(field, csvValue));

      case 'TaxonomyFieldType':
      case 'TaxonomyFieldTypeMulti':
        // Handled separately after item create/update
        return Promise.resolve(undefined);

      default:
        return Promise.resolve(undefined);
    }
  }

  // --- Choice / MultiChoice ---

  private _convertChoice(
    field: IListField, csvValue: string, allowFillIn: boolean
  ): Promise<IFieldValueResult | undefined> {
    if (!allowFillIn && field.choices && field.choices.indexOf(csvValue) < 0) {
      return Promise.reject(new Error(
        formatString(strings.ErrorChoiceValueInvalid,
          csvValue, field.displayName || field.internalName)
      ));
    }
    return Promise.resolve({
      fieldName: field.internalName, value: csvValue
    });
  }

  private _convertMultiChoice(
    field: IListField, csvValue: string, allowFillIn: boolean
  ): Promise<IFieldValueResult | undefined> {
    const choices: string[] = csvValue.split(';').map((s: string) => s.trim());
    if (!allowFillIn && field.choices) {
      const invalidChoices: string[] = choices.filter(
        (c: string) => field.choices.indexOf(c) < 0
      );
      if (invalidChoices.length > 0) {
        return Promise.reject(new Error(
          formatString(strings.ErrorMultiChoiceValuesInvalid,
            invalidChoices.join(', '),
            field.displayName || field.internalName)
        ));
      }
    }
    return Promise.resolve({
      fieldName: field.internalName,
      value: { results: choices }
    });
  }

  // --- Lookup ---

  private _convertLookup(
    webUrl: string, field: IListField, csvValue: string
  ): Promise<IFieldValueResult | undefined> {
    if (!field.lookupListId) {
      return Promise.resolve(undefined);
    }
    return this._service.resolveLookupValue(
      webUrl, field.lookupListId,
      field.lookupFieldName || 'Title', csvValue
    ).then((id: number | undefined) => {
      if (id !== undefined) {
        return {
          fieldName: field.internalName + 'Id',
          value: id
        };
      }
      throw new Error(
        formatString(strings.ErrorLookupValueNotFound,
          csvValue, field.displayName || field.internalName)
      );
    });
  }

  private _convertLookupMulti(
    webUrl: string, field: IListField, csvValue: string
  ): Promise<IFieldValueResult | undefined> {
    if (!field.lookupListId) {
      return Promise.resolve(undefined);
    }
    const lookupVals: string[] = csvValue.split(';').map((s: string) => s.trim());
    // Resolve each value individually to identify which ones fail
    const lookupPromises: Promise<{ val: string; id: number | undefined }>[] =
      lookupVals.map((v: string) =>
        this._service.resolveLookupValue(
          webUrl, field.lookupListId, field.lookupFieldName || 'Title', v
        ).then((id: number | undefined) => ({ val: v, id: id }))
      );
    return Promise.all(lookupPromises).then(
      (results: { val: string; id: number | undefined }[]) => {
        const resolved: number[] = results
          .filter((r: { val: string; id: number | undefined }) => r.id !== undefined)
          .map((r: { val: string; id: number | undefined }) => r.id as number);
        const unresolved: string[] = results
          .filter((r: { val: string; id: number | undefined }) => r.id === undefined)
          .map((r: { val: string; id: number | undefined }) => r.val);
        if (unresolved.length > 0) {
          throw new Error(
            formatString(strings.ErrorLookupValuesNotFound,
              unresolved.join(', '),
              field.displayName || field.internalName)
          );
        }
        return {
          fieldName: field.internalName + 'Id',
          value: { results: resolved }
        };
      }
    );
  }

  // --- User ---

  private _convertUser(
    webUrl: string, field: IListField, csvValue: string
  ): Promise<IFieldValueResult | undefined> {
    return this._service.resolveUser(webUrl, csvValue)
      .then((id: number | undefined) => {
        if (id !== undefined) {
          return {
            fieldName: field.internalName + 'Id',
            value: id
          };
        }
        throw new Error(
          formatString(strings.ErrorUserNotFound,
            csvValue, field.displayName || field.internalName)
        );
      });
  }

  private _convertUserMulti(
    webUrl: string, field: IListField, csvValue: string
  ): Promise<IFieldValueResult | undefined> {
    const userNames: string[] = csvValue.split(';').map((s: string) => s.trim());
    // Resolve each user individually to identify which ones fail
    const userResolvePromises: Promise<{ name: string; id: number | undefined }>[] =
      userNames.map((name: string) =>
        this._service.resolveUser(webUrl, name)
          .then((id: number | undefined) => ({ name: name, id: id }))
      );
    return Promise.all(userResolvePromises).then(
      (results: { name: string; id: number | undefined }[]) => {
        const resolvedIds: number[] = results
          .filter((r: { name: string; id: number | undefined }) => r.id !== undefined)
          .map((r: { name: string; id: number | undefined }) => r.id as number);
        const unresolvedNames: string[] = results
          .filter((r: { name: string; id: number | undefined }) => r.id === undefined)
          .map((r: { name: string; id: number | undefined }) => r.name);
        if (unresolvedNames.length > 0) {
          throw new Error(
            formatString(strings.ErrorUsersNotFound,
              unresolvedNames.join(', '),
              field.displayName || field.internalName)
          );
        }
        return {
          fieldName: field.internalName + 'Id',
          value: { results: resolvedIds }
        };
      }
    );
  }

  // --- URL ---

  private _convertUrl(field: IListField, csvValue: string): IFieldValueResult {
    const parts: string[] = csvValue.split(',');
    let urlVal: string = parts[0].trim();
    const desc: string = parts.length > 1
      ? parts.slice(1).join(',').trim()
      : '';

    // Auto-prefix protocol if missing
    if (urlVal && !/^[a-zA-Z][a-zA-Z0-9+\-.]*:\/\//i.test(urlVal)) {
      urlVal = 'https://' + urlVal;
    }

    // If no explicit description was given, derive one from the URL
    let displayDesc: string = desc;
    if (!displayDesc) {
      const match: RegExpMatchArray =
        urlVal.match(/^(?:https?:\/\/)?([^\/\?#]+)/i);
      displayDesc = match ? match[1] : urlVal;
    }

    return {
      fieldName: field.internalName,
      value: { Url: urlVal, Description: displayDesc }
    };
  }

  // --- Primitive parsers ---

  private _parseBooleanValue(value: string): boolean {
    const lower: string = value.toLowerCase().trim();
    return lower === 'true' || lower === '1' ||
      lower === 'ja' || lower === 'yes' || lower === 'wahr';
  }

  /**
   * Parse a date in German (DD.MM.YYYY), US (MM/DD/YYYY) or ISO format.
   * Specific national formats are checked first so they are not
   * mis-parsed by the more permissive `new Date(value)` fallback.
   */
  private _parseDateValue(value: string): string | undefined {
    if (!value || value.trim() === '') {
      return undefined;
    }

    let date: Date;

    // German format: DD.MM.YYYY or DD.MM.YYYY HH:MM:SS
    const germanMatch: RegExpMatchArray | undefined = value.match(
      /^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
    );
    if (germanMatch) {
      const day: number = parseInt(germanMatch[1], 10);
      const month: number = parseInt(germanMatch[2], 10) - 1;
      const year: number = parseInt(germanMatch[3], 10);
      const hours: number = germanMatch[4] ? parseInt(germanMatch[4], 10) : 0;
      const minutes: number = germanMatch[5] ? parseInt(germanMatch[5], 10) : 0;
      const seconds: number = germanMatch[6] ? parseInt(germanMatch[6], 10) : 0;
      date = new Date(year, month, day, hours, minutes, seconds);
      if (!isNaN(date.getTime())) {
        return date.toISOString();
      }
    }

    // US format: MM/DD/YYYY — checked before the generic parse so
    // browsers that interpret slashes ambiguously can't get it wrong.
    const usMatch: RegExpMatchArray | undefined =
      value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (usMatch) {
      date = new Date(
        parseInt(usMatch[3], 10),
        parseInt(usMatch[1], 10) - 1,
        parseInt(usMatch[2], 10)
      );
      if (!isNaN(date.getTime())) {
        return date.toISOString();
      }
    }

    // ISO format or anything else the browser understands
    date = new Date(value);
    if (!isNaN(date.getTime())) {
      return date.toISOString();
    }

    return undefined;
  }
}
