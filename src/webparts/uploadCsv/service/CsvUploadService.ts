// tslint:disable:no-any
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { sp, Web, SearchResults } from '@pnp/sp';
import { taxonomy, ITermStore } from '@pnp/sp-taxonomy';
import {
  ISiteCollection,
  IWeb,
  IListInfo,
  IListField,
  SpFieldType
} from '../models';
import { LOG } from '../utils/log';

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
        'Choices', 'LookupList', 'LookupField',
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
            termSetId: f.TermSetId || undefined,
            sspId: f.SspId || undefined,
            defaultValue: f.DefaultValue || undefined
          });
        });
        return this._enrichTaxonomyFields(webUrl, listId, fields);
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
    const web: Web = new Web(webUrl);
    return web.lists.getByTitle('TaxonomyHiddenList').items
      .filter(
        "Title eq '" + this._escapeODataValue(label) + "'"
      )
      .select('Id', 'Title', 'IdForTerm')
      .top(1)
      .get()
      .then((items: any[]) => {
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
    const store: ITermStore = sspId
      ? taxonomy.termStores.getById(sspId)
      : taxonomy.getDefaultSiteCollectionTermStore();
    return store.getTermSetById(termSetId).terms.get()
      .then((terms: any[]) => {
        let match: any = undefined;
        for (let i: number = 0; i < terms.length; i++) {
          if (terms[i].Name === label) {
            match = terms[i];
            break;
          }
        }
        if (match) {
          const rawGuid: string = match.Id || '';
          // Normalise GUID: strip /Guid(...)/ wrapper and curly braces
          const cleanGuid: string = rawGuid
            .replace(/^\/Guid\((.*)\)\/$/i, '$1')
            .replace(/^\{|\}$/g, '');
          return {
            wssId: -1,
            label: match.Name || label,
            termGuid: cleanGuid
          };
        }
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
    // tslint:disable-next-line:no-any
    const payload: { [key: string]: any } = {};
    payload[fieldInternalName] = {
      __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
      Label: label,
      TermGuid: termGuid,
      WssId: -1
    };
    const web: Web = new Web(webUrl);
    return web.lists.getById(listId).items
      .getById(itemId)
      .update(payload)
      .then(() => { /* void */ })
      .catch((err: any) => {
        console.error(LOG, 'setTaxonomyFieldValue ERROR — itemId:', itemId,
          'field:', fieldInternalName,
          'label:', label, 'termGuid:', termGuid,
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
    const taxValue: string = values.map(
      (v: { wssId: number; label: string; termGuid: string }) =>
        '-1;#' + v.label + '|' + v.termGuid
    ).join(';#');
    // tslint:disable-next-line:no-any
    const payload: { [key: string]: any } = {};
    payload[hiddenFieldName] = taxValue;
    const web: Web = new Web(webUrl);
    return web.lists.getById(listId).items
      .getById(itemId)
      .update(payload)
      .then(() => { /* void */ })
      .catch((err: any) => {
        console.error(LOG, 'setTaxonomyMultiFieldValue ERROR — itemId:', itemId,
          'field:', fieldInternalName,
          'hiddenField:', hiddenFieldName,
          'taxValue:', taxValue,
          err && err.data ? JSON.stringify(err.data) : '',
          err && err.status ? 'status=' + err.status : '',
          err.message || err);
        throw err;
      });
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
