/**
 * SiteSearchService.ts
 *
 * Service class that wraps the Microsoft Graph Search API to discover SharePoint
 * sites accessible to the currently signed-in user.  Used by the CustomNav
 * application customizer to populate the navigation menu.
 *
 * All calls are made through MSGraphClientV3, which the SPFx framework provides
 * pre-authenticated with the user's delegated permissions.
 */
import { MSGraphClientV3 } from '@microsoft/sp-http';

/**
 * Represents a single SharePoint site returned from the Graph Search API.
 * Only the fields explicitly requested in the `fields` array are populated;
 * all others will be undefined.
 */
export interface ISiteHit {
  /** The unique identifier of the site resource (e.g. "tenant,site-guid,web-guid"). */
  id: string;

  /** Absolute URL of the SharePoint site (e.g. https://tenant.sharepoint.com/sites/MyHub). */
  webUrl?: string;

  /** Human-readable site title as configured in the SharePoint admin centre. */
  displayName?: string;

  /** The site's short URL segment name, useful as a fallback display label. */
  name?: string;

  /** Optional description set by the site owner. */
  description?: string;
}

/**
 * SiteSearchService
 *
 * Provides two public methods for site discovery:
 *  - `searchSites`         — paginated keyword search across accessible sites.
 *  - `listAccessibleSites` — aggregates all accessible sites up to a configurable
 *                             maximum using `searchSites` with a wildcard query.
 */
export default class SiteSearchService {
  /**
   * Constructs a new SiteSearchService.
   *
   * @param graphClient — An authenticated MSGraphClientV3 instance obtained from
   *                      `this.context.msGraphClientFactory.getClient('3')` inside
   *                      an SPFx extension or web part.
   */
  constructor(private graphClient: MSGraphClientV3) {}

  /**
   * Searches SharePoint sites accessible to the current user using the
   * Microsoft Graph Search API (`POST /search/query`).
   *
   * When `queryString` is empty or whitespace a wildcard ('*') is used so that
   * all accessible sites are returned (subject to Graph search indexing).
   *
   * @param queryString - Keyword(s) to search for in site names/descriptions.
   *                      Pass an empty string or '*' for an unfiltered listing.
   * @param from        - Zero-based offset for pagination (default 0).
   * @param size        - Number of results to return per page (default 25).
   * @returns A page of site results with a `more` flag and optional `total` count.
   */
  public async searchSites(queryString: string, from = 0, size = 25): Promise<{ items: ISiteHit[]; more?: boolean; total?: number; }> {
    // Default to wildcard when no meaningful query text is provided.
    const q = queryString?.trim() ? queryString.trim() : '*';

    const body = {
      requests: [
        {
          entityTypes: ['site'],
          query: { queryString: q },
          from,
          size,
          fields: ['id', 'webUrl', 'displayName', 'name', 'description']
        }
      ]
    };

    const res = await this.graphClient
      .api('/search/query')
      .version('v1.0')
      .post(body); // POST /search/query 【3-76c44f】

    const container = res?.value?.[0]?.hitsContainers?.[0];
    const hits = container?.hits ?? [];

    const items = hits.map((h: any) => h?.resource).filter(Boolean) as ISiteHit[];

    return {
      items,
      more: container?.moreResultsAvailable,
      total: container?.total
    };
  }

  /**
   * Retrieves all SharePoint sites accessible to the current user by issuing
   * sequential wildcard search pages until all results have been collected or
   * `max` items have been accumulated.
   *
   * Results are deduplicated by `webUrl` (or `id` as fallback) before returning
   * to guard against the same site appearing across multiple search pages.
   *
   * @param max - Upper bound on the number of sites to return (default 300).
   *              Prevents unbounded loop execution on very large tenants.
   * @returns Deduplicated array of all accessible ISiteHit objects.
   */
  public async listAccessibleSites(max = 300): Promise<ISiteHit[]> {
    const pageSize = 100;           // Items per page (safe below the Graph API limit).
    const collected: ISiteHit[] = [];
    let from = 0;
    let more = true;               // Assume at least one page exists; updated after each call.

    // Continue fetching pages until Graph reports no more results or we hit the max cap.
    while (more && collected.length < max) {
      const page = await this.searchSites('*', from, pageSize);
      const items = page.items ?? [];

      collected.push(...items);
      from += pageSize;
      // Stop if the last page was empty or Graph signals no further results.
      more = Boolean(page.more) && items.length > 0;
    }

    // Deduplicate using a Map keyed by lowercased webUrl (or id as fallback) to handle
    // cases where the same site appears in overlapping search result pages.
    const deduped = new Map<string, ISiteHit>();
    for (const site of collected) {
      const key = (site.webUrl || site.id || '').toLowerCase();
      if (!key || deduped.has(key)) {
        continue;  // Skip entries with no usable key, or already seen.
      }

      deduped.set(key, site);
    }

    return Array.from(deduped.values());
  }
}