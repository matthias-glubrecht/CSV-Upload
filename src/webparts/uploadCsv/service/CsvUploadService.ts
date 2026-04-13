// tslint:disable:no-any
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { sp, Web, SearchResults } from '@pnp/sp';
import { taxonomy, ITermStore } from '@pnp/sp-taxonomy';
import * as strings from 'UploadCsvWebPartStrings';
import {
  ISiteCollection,
  IWeb,
  IListInfo,
  IListField,
  SpFieldType
} from '../models';

const LOG: string = '[CsvUpload]';

/**
 * Service class encapsulating all SharePoint data access
 * for the CSV Upload web part. Uses @pnp/sp for data access.
 */
export default class CsvUploadService {
  private _context: WebPartContext;

  constructor(context: WebPartContext) {
    this._context = context;
    sp.setup({ spfxContext: context });
  }

  // ─── Site Collections ──────────────────────────────────────────

  /**
   * Search for site collections using the PnP search API.
   */
  public searchSiteCollections(
    query: string
  ): Promise<ISiteCollection[]> {
    const searchQuery: string = query
      ? 'contentclass:STS_Site AND Title:' + query + '*'
      : 'contentclass:STS_Site';

    return sp.search({
      Querytext: searchQuery,
      SelectProperties: ['Title', 'SPSiteUrl'],
      RowLimit: 50,
      TrimDuplicates: true
    }).then((searchResults: SearchResults) => {
      const results: ISiteCollection[] = [];
      searchResults.PrimarySearchResults.forEach(
        (row: any) => {
          const siteUrl: string = row.SPSiteUrl || '';
          if (siteUrl) {
            results.push({
              title: (row.Title as string) || siteUrl,
              url: siteUrl
            });
          }
        }
      );
      return this._deduplicateSites(results);
    });
  }

  /**
   * Get the current site collection info.
   */
  public getCurrentSiteCollection(): ISiteCollection {
    const siteUrl: string =
      this._context.pageContext.site.absoluteUrl;
    const segments: string[] = siteUrl.split('/').filter(
      (s: string) => s.length > 0
    );
    return {
      title: segments.pop() || 'Root',
      url: siteUrl
    };
  }

  // ─── Webs ──────────────────────────────────────────────────────

  /**
   * Get all webs (sub-sites) within a site collection.
   */
  public getWebs(
    siteCollectionUrl: string
  ): Promise<IWeb[]> {
    const rootWeb: Web = new Web(siteCollectionUrl);
    return rootWeb.webs
      .select('Title', 'Url', 'Id')
      .get()
      .then((data: any[]) => {
        const webs: IWeb[] = [];
        webs.push({
          title: 'Root',
          url: siteCollectionUrl,
          id: ''
        });
        data.forEach((w: any) => {
          webs.push({
            title: w.Title,
            url: w.Url,
            id: w.Id
          });
        });
        return webs;
      });
  }

  /**
   * Get the current web info.
   */
  public getCurrentWeb(): IWeb {
    return {
      title: this._context.pageContext.web.title,
      url: this._context.pageContext.web.absoluteUrl,
      id: this._context.pageContext.web.id.toString()
    };
  }

  // ─── Lists ─────────────────────────────────────────────────────

  /**
   * Get all custom lists (no document libraries) from a web.
   */
  public getLists(webUrl: string): Promise<IListInfo[]> {
    const web: Web = new Web(webUrl);
    return web.lists
      .filter('BaseTemplate eq 100 and Hidden eq false')
      .select('Title', 'Id')
      .get()
      .then((data: any[]) => {
        return data.map((list: any) => ({
          title: list.Title as string,
          id: list.Id as string
        }));
      });
  }

  // ─── Fields ────────────────────────────────────────────────────

  /**
   * Get relevant (user-facing) fields for a list.
   */
  public getListFields(
    webUrl: string, listId: string
  ): Promise<IListField[]> {
    const web: Web = new Web(webUrl);
    return web.lists.getById(listId).fields
      .filter('Hidden eq false and ReadOnlyField eq false')
      .select(
        'InternalName', 'Title', 'TypeAsString', 'Required',
        'Choices', 'LookupList', 'LookupField', 'LookupWebId',
        'TermSetId', 'SspId', 'DefaultValue'
      )
      .get()
      .then((data: any[]) => {
        const fields: IListField[] = [];
        const skipFields: string[] = [
          'ContentType', 'Attachments',
          '_ModerationComments', 'Edit',
          'LinkTitleNoMenu', 'LinkTitle', 'DocIcon',
          'ItemChildCount', 'FolderChildCount',
          '_ComplianceFlags', '_ComplianceTag',
          '_ComplianceTagWrittenTime',
          '_ComplianceTagUserId',
          'AppAuthor', 'AppEditor', 'ComplianceAssetId'
        ];
        data.forEach((f: any) => {
          if (skipFields.indexOf(f.InternalName) >= 0) {
            return;
          }
          const fieldType: SpFieldType =
            this._mapFieldType(f.TypeAsString);
          if (fieldType === 'Unknown') {
            return;
          }
          fields.push({
            internalName: f.InternalName,
            displayName: f.Title,
            fieldType: fieldType,
            required: f.Required || false,
            choices: f.Choices
              ? f.Choices.results || f.Choices
              : undefined,
            lookupListId: f.LookupList || undefined,
            lookupFieldName: f.LookupField || undefined,
            lookupWebId: f.LookupWebId || undefined,
            termSetId: f.TermSetId || undefined,
            sspId: f.SspId || undefined,
            defaultValue: f.DefaultValue || undefined
          });
        });
        return this._enrichTaxonomyFields(
          webUrl, listId, fields
        ).then((enriched: IListField[]) => {
          console.log(LOG, 'Fields loaded:', enriched.length,
            enriched.map((f: IListField) => ({
              name: f.internalName, type: f.fieldType,
              termSetId: f.termSetId, sspId: f.sspId,
              taxHidden: f.taxonomyHiddenFieldName
            }))
          );
          return enriched;
        });
      });
  }

  // ─── Lookup Resolution ─────────────────────────────────────────

  /**
   * Resolve a lookup display value to the lookup item ID.
   */
  public resolveLookupValue(
    webUrl: string,
    lookupListId: string,
    lookupFieldName: string,
    value: string
  ): Promise<number | undefined> {
    const filterField: string = lookupFieldName || 'Title';
    const web: Web = new Web(webUrl);
    return web.lists.getById(lookupListId).items
      .filter(
        filterField + " eq '" +
        this._escapeODataValue(value) + "'"
      )
      .select('Id')
      .top(1)
      .get()
      .then((items: any[]) => {
        if (items.length > 0) {
          return items[0].Id as number;
        }
        return undefined;
      });
  }

  /**
   * Resolve multiple lookup values (for LookupMulti fields).
   */
  public resolveLookupValues(
    webUrl: string,
    lookupListId: string,
    lookupFieldName: string,
    values: string[]
  ): Promise<number[]> {
    const promises: Promise<number | undefined>[] = values.map(
      (v: string) => this.resolveLookupValue(
        webUrl, lookupListId, lookupFieldName, v.trim()
      )
    );
    return Promise.all(promises).then(
      (results: (number | undefined)[]) => {
        return results.filter(
          (r: number | undefined) => r !== undefined
        ) as number[];
      }
    );
  }

  // ─── User Resolution ───────────────────────────────────────────

  /**
   * Resolve a user display name or login name to a user ID.
   */
  public resolveUser(
    webUrl: string, userIdentifier: string
  ): Promise<number | undefined> {
    const web: Web = new Web(webUrl);
    return web.ensureUser(userIdentifier)
      .then((result: any) => {
        if (result && result.data && result.data.Id) {
          return result.data.Id as number;
        }
        return undefined;
      })
      .catch(() => {
        // If ensureUser fails, try searching by title
        return this._resolveUserByTitle(
          webUrl, userIdentifier
        );
      });
  }

  // ─── Taxonomy Resolution ───────────────────────────────────────

  /**
   * Resolve a taxonomy value by label.
   * First searches TaxonomyHiddenList (fast, works for terms
   * already used on this site). If not found there, falls back
   * to searching the TaxonomyHiddenList without a term-set
   * filter (the term may exist under a different term set in the
   * hidden list with a matching title).
   */
  public resolveTaxonomyValue(
    webUrl: string,
    termSetId: string,
    label: string
  ): Promise<{
    wssId: number; label: string; termGuid: string
  } | undefined> {
    console.log(LOG, 'resolveTaxonomyValue called — termSetId:', termSetId, 'label:', JSON.stringify(label));
    const web: Web = new Web(webUrl);
    return web.lists.getByTitle('TaxonomyHiddenList').items
      .filter(
        "Title eq '" + this._escapeODataValue(label) + "'"
      )
      .select('Id', 'Title', 'IdForTerm')
      .top(1)
      .get()
      .then((items: any[]) => {
        console.log(LOG, 'resolveTaxonomyValue result for',
          JSON.stringify(label), '— items returned:',
          items.length, items);
        if (items.length > 0) {
          return {
            wssId: items[0].Id,
            label: items[0].Title,
            termGuid: items[0].IdForTerm || ''
          };
        }
        return undefined;
      })
      .catch((err: Error) => {
        console.warn(LOG, 'resolveTaxonomyValue ERROR for', JSON.stringify(label), err);
        return undefined;
      });
  }

  /**
   * Search the term store for a term by label within a term set.
   * Uses @pnp/sp-taxonomy to get all terms from the term set
   * and finds the matching one by name.
   * Returns the term GUID if found (wssId will be -1 since
   * the term is not yet in TaxonomyHiddenList).
   */
  public resolveTermFromStore(
    termSetId: string,
    label: string,
    sspId?: string
  ): Promise<{
    wssId: number; label: string; termGuid: string
  } | undefined> {
    console.log(LOG, 'resolveTermFromStore called — termSetId:',
      termSetId, 'sspId:', sspId, 'label:', JSON.stringify(label));
    const store: ITermStore = sspId
      ? taxonomy.termStores.getById(sspId)
      : taxonomy.getDefaultSiteCollectionTermStore();
    return store.getTermSetById(termSetId).terms.get()
      .then((terms: any[]) => {
        console.log(LOG, 'resolveTermFromStore — terms in set:',
          terms.length,
          terms.map((t: any) => ({ Name: t.Name, Id: t.Id })));
        // Log raw first term object so we can see the exact GUID format
        if (terms.length > 0) {
          console.log(LOG, 'resolveTermFromStore — RAW first term object:',
            JSON.stringify(terms[0]));
        }
        let match: any = undefined;
        for (let i: number = 0; i < terms.length; i++) {
          if (terms[i].Name === label) {
            match = terms[i];
            break;
          }
        }
        if (match) {
          // tslint:disable-next-line:no-any
          const rawGuid: string = match.Id || '';
          // Normalise GUID: strip /Guid(...)/ wrapper and curly braces
          const cleanGuid: string = rawGuid
            .replace(/^\/Guid\((.*)\)\/$/i, '$1')
            .replace(/^\{|\}$/g, '');
          console.log(LOG, 'resolveTermFromStore found:',
            match.Name,
            'rawGuid:', JSON.stringify(rawGuid),
            'cleanGuid:', JSON.stringify(cleanGuid));
          return {
            wssId: -1,
            label: match.Name || label,
            termGuid: cleanGuid
          };
        }
        console.warn(LOG, 'resolveTermFromStore — term not found for label:',
          JSON.stringify(label),
          '— available names:',
          terms.map((t: any) => t.Name));
        return undefined;
      })
      .catch((err: Error) => {
        console.warn(LOG, 'resolveTermFromStore ERROR for',
          JSON.stringify(label), err);
        return undefined;
      });
  }

  // ─── Item CRUD ─────────────────────────────────────────────────

  /**
   * Get existing items from the list for upsert logic.
   */
  public getExistingItems(
    webUrl: string,
    listId: string,
    keyFieldInternalName: string,
    keyValues: string[]
  ): Promise<{ [key: string]: number }> {
    const result: { [key: string]: number } = {};
    const batchSize: number = 15;
    const batches: string[][] = [];

    for (let i: number = 0; i < keyValues.length; i += batchSize) {
      batches.push(keyValues.slice(i, i + batchSize));
    }

    const fetchBatch: (idx: number) => Promise<void> =
      (idx: number) => {
        if (idx >= batches.length) {
          return Promise.resolve();
        }
        const batch: string[] = batches[idx];
        const filterParts: string[] = batch.map(
          (v: string) => keyFieldInternalName +
            " eq '" + this._escapeODataValue(v) + "'"
        );
        const filterQuery: string = filterParts.join(' or ');
        const web: Web = new Web(webUrl);
        return web.lists.getById(listId).items
          .filter(filterQuery)
          .select('Id', keyFieldInternalName)
          .top(5000)
          .get()
          .then((items: any[]) => {
            items.forEach((item: any) => {
              const keyVal: string =
                String(item[keyFieldInternalName] || '');
              result[keyVal] = item.Id;
            });
            return fetchBatch(idx + 1);
          });
      };

    return fetchBatch(0).then(() => result);
  }

  /**
   * Create a new list item.
   */
  public createItem(
    webUrl: string, listId: string, fieldValues: any
  ): Promise<{ Id: number }> {
    const web: Web = new Web(webUrl);
    return web.lists.getById(listId).items
      .add(fieldValues)
      .then((result: any) => result.data as { Id: number });
  }

  /**
   * Update an existing list item.
   */
  public updateItem(
    webUrl: string,
    listId: string,
    itemId: number,
    fieldValues: any
  ): Promise<void> {
    const web: Web = new Web(webUrl);
    return web.lists.getById(listId).items
      .getById(itemId)
      .update(fieldValues)
      .then(() => { /* void */ });
  }

  /**
   * Set a single-value taxonomy field on an item.
   * Uses the SP.Taxonomy.TaxonomyFieldValue metadata object
   * on the field's own internal name.
   */
  public setTaxonomyFieldValue(
    webUrl: string,
    listId: string,
    itemId: number,
    fieldInternalName: string,
    _hiddenFieldName: string,
    label: string,
    termGuid: string
  ): Promise<void> {
    console.log(LOG, 'setTaxonomyFieldValue — itemId:', itemId,
      'field:', fieldInternalName,
      'label:', label, 'termGuid:', termGuid);
    // tslint:disable-next-line:no-any
    const payload: { [key: string]: any } = {};
    payload[fieldInternalName] = {
      __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
      Label: label,
      TermGuid: termGuid,
      WssId: -1
    };
    console.log(LOG, 'setTaxonomyFieldValue PAYLOAD:',
      JSON.stringify(payload, undefined, 2));
    const web: Web = new Web(webUrl);
    return web.lists.getById(listId).items
      .getById(itemId)
      .update(payload)
      .then(() => {
        console.log(LOG, 'setTaxonomyFieldValue REST call OK — itemId:', itemId,
          'field:', fieldInternalName);
        // Read back the field to verify the value was actually persisted
        return web.lists.getById(listId).items
          .getById(itemId)
          .select(fieldInternalName)
          .get();
      })
      .then((readBack: any) => {
        console.log(LOG, 'setTaxonomyFieldValue READ-BACK — itemId:', itemId,
          'field:', fieldInternalName,
          'value:', JSON.stringify(readBack[fieldInternalName]));
        // tslint:disable-next-line:no-any
        const val: any = readBack[fieldInternalName];
        if (!val || (!val.TermGuid && !val.Label)) {
          console.error(LOG,
            'setTaxonomyFieldValue VERIFICATION FAILED — field appears empty after update!',
            'itemId:', itemId, 'field:', fieldInternalName,
            'full readBack:', JSON.stringify(readBack));
        }
      })
      .catch((err: any) => {
        console.error(LOG, 'setTaxonomyFieldValue ERROR — itemId:', itemId,
          'field:', fieldInternalName,
          'label:', label, 'termGuid:', termGuid);
        console.error(LOG, 'setTaxonomyFieldValue ERROR details:',
          err && err.data ? JSON.stringify(err.data) : '',
          err && err.status ? 'status=' + err.status : '',
          err.message || err);
        throw err;
      });
  }

  /**
   * Set a multi-value taxonomy field on an item.
   * Writes to the hidden note field associated with the
   * taxonomy column.
   */
  public setTaxonomyMultiFieldValue(
    webUrl: string,
    listId: string,
    itemId: number,
    fieldInternalName: string,
    hiddenFieldName: string,
    values: Array<{
      wssId: number; label: string; termGuid: string
    }>
  ): Promise<void> {
    console.log(LOG, 'setTaxonomyMultiFieldValue — input values:',
      JSON.stringify(values));
    const taxValue: string = values.map(
      (v: { wssId: number; label: string; termGuid: string }) =>
        '-1;#' + v.label + '|' + v.termGuid
    ).join(';#');
    console.log(LOG, 'setTaxonomyMultiFieldValue — itemId:', itemId,
      'field:', fieldInternalName, 'hiddenField:', hiddenFieldName,
      'taxValue:', taxValue);
    // tslint:disable-next-line:no-any
    const payload: { [key: string]: any } = {};
    payload[hiddenFieldName] = taxValue;
    console.log(LOG, 'setTaxonomyMultiFieldValue PAYLOAD:',
      JSON.stringify(payload, undefined, 2));
    const web: Web = new Web(webUrl);
    return web.lists.getById(listId).items
      .getById(itemId)
      .update(payload)
      .then(() => {
        console.log(LOG, 'setTaxonomyMultiFieldValue REST call OK — itemId:', itemId,
          'field:', fieldInternalName);
        // Read back the taxonomy field to verify the value was persisted
        return web.lists.getById(listId).items
          .getById(itemId)
          .select(fieldInternalName)
          .get();
      })
      .then((readBack: any) => {
        console.log(LOG, 'setTaxonomyMultiFieldValue READ-BACK — itemId:', itemId,
          'field:', fieldInternalName,
          'value:', JSON.stringify(readBack[fieldInternalName]));
      })
      .catch((err: any) => {
        console.error(LOG, 'setTaxonomyMultiFieldValue ERROR — itemId:', itemId,
          'field:', fieldInternalName,
          'hiddenField:', hiddenFieldName,
          'taxValue:', taxValue);
        console.error(LOG, 'setTaxonomyMultiFieldValue ERROR details:',
          err && err.data ? JSON.stringify(err.data) : '',
          err && err.status ? 'status=' + err.status : '',
          err.message || err);
        throw err;
      });
  }

  // ─── Field Value Conversion ────────────────────────────────────

  /**
   * Convert a CSV string value to the appropriate field value
   * for the SharePoint REST API.
   */
  public convertFieldValue(
    webUrl: string,
    field: IListField,
    csvValue: string,
    allowChoiceFillIn: boolean
  ): Promise<{ fieldName: string; value: any } | undefined> {

    if (csvValue === undefined || csvValue === '') {
      return Promise.resolve(undefined);
    }

    switch (field.fieldType) {
      case 'Text':
        return Promise.resolve({
          fieldName: field.internalName, value: csvValue
        });

      case 'Note':
        return Promise.resolve({
          fieldName: field.internalName, value: csvValue
        });

      case 'Number':
      case 'Currency': {
        const numVal: number =
          parseFloat(csvValue.replace(',', '.'));
        if (isNaN(numVal)) {
          return Promise.resolve(undefined);
        }
        return Promise.resolve({
          fieldName: field.internalName, value: numVal
        });
      }

      case 'Boolean': {
        const boolVal: boolean =
          this._parseBooleanValue(csvValue);
        return Promise.resolve({
          fieldName: field.internalName, value: boolVal
        });
      }

      case 'DateTime': {
        const dateVal: string | undefined =
          this._parseDateValue(csvValue);
        if (!dateVal) {
          return Promise.resolve(undefined);
        }
        return Promise.resolve({
          fieldName: field.internalName, value: dateVal
        });
      }

      case 'Choice': {
        if (!allowChoiceFillIn && field.choices &&
            field.choices.indexOf(csvValue) < 0) {
          return Promise.reject(new Error(
            strings.ErrorChoiceValueInvalid
              .replace('{0}', csvValue)
              .replace('{1}', field.displayName || field.internalName)
          ));
        }
        return Promise.resolve({
          fieldName: field.internalName, value: csvValue
        });
      }

      case 'MultiChoice': {
        const choices: string[] =
          csvValue.split(';').map((s: string) => s.trim());
        if (!allowChoiceFillIn && field.choices) {
          const invalidChoices: string[] = choices.filter(
            (c: string) => field.choices.indexOf(c) < 0
          );
          if (invalidChoices.length > 0) {
            return Promise.reject(new Error(
              strings.ErrorMultiChoiceValuesInvalid
                .replace('{0}', invalidChoices.join(', '))
                .replace('{1}', field.displayName || field.internalName)
            ));
          }
        }
        return Promise.resolve({
          fieldName: field.internalName,
          value: { results: choices }
        });
      }

      case 'Lookup': {
        if (!field.lookupListId) {
          return Promise.resolve(undefined);
        }
        return this.resolveLookupValue(
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
            strings.ErrorLookupValueNotFound
              .replace('{0}', csvValue)
              .replace('{1}', field.displayName || field.internalName)
          );
        });
      }

      case 'LookupMulti': {
        if (!field.lookupListId) {
          return Promise.resolve(undefined);
        }
        const lookupVals: string[] =
          csvValue.split(';').map((s: string) => s.trim());
        // Resolve each value individually to identify which ones fail
        const lookupPromises: Promise<{ val: string; id: number | undefined }>[] =
          lookupVals.map((v: string) =>
            this.resolveLookupValue(
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
                strings.ErrorLookupValuesNotFound
                  .replace('{0}', unresolved.join(', '))
                  .replace('{1}', field.displayName || field.internalName)
              );
            }
            return {
              fieldName: field.internalName + 'Id',
              value: { results: resolved }
            };
          }
        );
      }

      case 'User': {
        return this.resolveUser(webUrl, csvValue)
          .then((id: number | undefined) => {
            if (id !== undefined) {
              return {
                fieldName: field.internalName + 'Id',
                value: id
              };
            }
            throw new Error(
              strings.ErrorUserNotFound
                .replace('{0}', csvValue)
                .replace('{1}', field.displayName || field.internalName)
            );
          });
      }

      case 'UserMulti': {
        const userNames: string[] =
          csvValue.split(';').map((s: string) => s.trim());
        // Resolve each user individually to identify which ones fail
        const userResolvePromises: Promise<{ name: string; id: number | undefined }>[] =
          userNames.map(
            (name: string) => this.resolveUser(webUrl, name)
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
                strings.ErrorUsersNotFound
                  .replace('{0}', unresolvedNames.join(', '))
                  .replace('{1}', field.displayName || field.internalName)
              );
            }
            return {
              fieldName: field.internalName + 'Id',
              value: { results: resolvedIds }
            };
          }
        );
      }

      case 'URL': {
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
          try {
            // Extract hostname as a readable description
            const match: RegExpMatchArray =
              urlVal.match(/^(?:https?:\/\/)?([^\/\?#]+)/i);
            displayDesc = match ? match[1] : urlVal;
          } catch (e) {
            displayDesc = urlVal;
          }
        }

        return Promise.resolve({
          fieldName: field.internalName,
          value: { Url: urlVal, Description: displayDesc }
        });
      }

      case 'TaxonomyFieldType':
      case 'TaxonomyFieldTypeMulti':
        // Handled separately after item creation/update
        console.log(LOG, 'convertFieldValue — skipping taxonomy field',
          field.internalName, '(handled post-create)');
        return Promise.resolve(undefined);

      default:
        return Promise.resolve(undefined);
    }
  }

  // ─── Private Helpers ───────────────────────────────────────────

  private _enrichTaxonomyFields(
    webUrl: string,
    listId: string,
    fields: IListField[]
  ): Promise<IListField[]> {
    const hasTaxonomy: boolean = fields.some(
      (f: IListField) =>
        f.fieldType === 'TaxonomyFieldType' ||
        f.fieldType === 'TaxonomyFieldTypeMulti'
    );
    if (!hasTaxonomy) {
      return Promise.resolve(fields);
    }
    return this._findTaxonomyHiddenFields(
      webUrl, listId, fields
    );
  }

  private _findTaxonomyHiddenFields(
    webUrl: string,
    listId: string,
    fields: IListField[]
  ): Promise<IListField[]> {
    const web: Web = new Web(webUrl);
    return web.lists.getById(listId).fields
      .select('InternalName', 'Id', 'TypeAsString', 'TextField')
      .get()
      .then((data: any[]) => {
        const taxToNoteGuid: { [name: string]: string } = {};
        const guidToName: { [guid: string]: string } = {};

        data.forEach((f: any) => {
          guidToName[f.Id.toLowerCase()] = f.InternalName;
          if ((f.TypeAsString === 'TaxonomyFieldType' ||
               f.TypeAsString === 'TaxonomyFieldTypeMulti') &&
              f.TextField) {
            taxToNoteGuid[f.InternalName] =
              f.TextField.toLowerCase();
          }
        });

        fields.forEach((field: IListField) => {
          if (field.fieldType === 'TaxonomyFieldType' ||
              field.fieldType === 'TaxonomyFieldTypeMulti') {
            const noteGuid: string =
              taxToNoteGuid[field.internalName];
            if (noteGuid) {
              const noteName: string = guidToName[noteGuid];
              if (noteName) {
                field.taxonomyHiddenFieldName = noteName;
              }
            }
            console.log(LOG, 'Taxonomy field enriched:',
              field.internalName, '— termSetId:', field.termSetId,
              'hiddenField:', field.taxonomyHiddenFieldName || '(none)');
          }
        });
        return fields;
      });
  }

  private _mapFieldType(typeAsString: string): SpFieldType {
    const map: { [key: string]: SpFieldType } = {
      'Text': 'Text',
      'Note': 'Note',
      'Number': 'Number',
      'Currency': 'Currency',
      'DateTime': 'DateTime',
      'Choice': 'Choice',
      'MultiChoice': 'MultiChoice',
      'Lookup': 'Lookup',
      'LookupMulti': 'LookupMulti',
      'User': 'User',
      'UserMulti': 'UserMulti',
      'Boolean': 'Boolean',
      'URL': 'URL',
      'TaxonomyFieldType': 'TaxonomyFieldType',
      'TaxonomyFieldTypeMulti': 'TaxonomyFieldTypeMulti'
    };
    return map[typeAsString] || 'Unknown';
  }

  private _resolveUserByTitle(
    webUrl: string, title: string
  ): Promise<number | undefined> {
    const web: Web = new Web(webUrl);
    return web.siteUsers
      .filter(
        "Title eq '" + this._escapeODataValue(title) + "'"
      )
      .select('Id')
      .top(1)
      .get()
      .then((users: any[]) => {
        if (users.length > 0) {
          return users[0].Id as number;
        }
        return undefined;
      });
  }

  private _parseBooleanValue(value: string): boolean {
    const lower: string = value.toLowerCase().trim();
    return lower === 'true' || lower === '1' ||
      lower === 'ja' || lower === 'yes' || lower === 'wahr';
  }

  private _parseDateValue(value: string): string | undefined {
    if (!value || value.trim() === '') {
      return undefined;
    }

    let date: Date;

    // German format: DD.MM.YYYY or DD.MM.YYYY HH:MM:SS
    const germanMatch: RegExpMatchArray | undefined =
      value.match(
        /^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
      );
    if (germanMatch) {
      const day: number = parseInt(germanMatch[1], 10);
      const month: number = parseInt(germanMatch[2], 10) - 1;
      const year: number = parseInt(germanMatch[3], 10);
      const hours: number = germanMatch[4]
        ? parseInt(germanMatch[4], 10) : 0;
      const minutes: number = germanMatch[5]
        ? parseInt(germanMatch[5], 10) : 0;
      const seconds: number = germanMatch[6]
        ? parseInt(germanMatch[6], 10) : 0;
      date = new Date(year, month, day, hours, minutes, seconds);
      if (!isNaN(date.getTime())) {
        return date.toISOString();
      }
    }

    // ISO format or US format
    date = new Date(value);
    if (!isNaN(date.getTime())) {
      return date.toISOString();
    }

    // MM/DD/YYYY
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

    return undefined;
  }

  private _escapeODataValue(value: string): string {
    return value.replace(/'/g, "''");
  }

  private _deduplicateSites(
    sites: ISiteCollection[]
  ): ISiteCollection[] {
    const seen: { [url: string]: boolean } = {};
    return sites.filter((site: ISiteCollection) => {
      const key: string = site.url.toLowerCase();
      if (seen[key]) {
        return false;
      }
      seen[key] = true;
      return true;
    });
  }
}
