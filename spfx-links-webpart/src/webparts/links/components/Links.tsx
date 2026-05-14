/**
 * Links.tsx  (spfx-links-webpart)
 *
 * React functional component that renders a flat ordered list of hyperlinks
 * fetched from a SharePoint list identified by GUID.  Items outside their
 * configured start/stop date window are filtered client-side before display.
 *
 * Features:
 *  - Uses SPHttpClient directly (no PnP SP dependency) for authentication.
 *  - Supports optional category filtering via an OData $filter clause.
 *  - Respects StartDate / StopDate columns for time-gated link visibility.
 *  - Configurable: open in new tab, show description, max items.
 */
import * as React from 'react';
import { useEffect, useState } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ILinksProps } from './ILinksProps';

/**
 * Raw shape of a link item row returned by the SharePoint REST API.
 * Field names match the internal column names configured in the list.
 */
interface ISharePointLinkItem {
  /** SharePoint item ID (auto-incremented integer). */
  Id: number;
  /** Display title of the link (used as anchor text). */
  Title: string;
  /** Hyperlink column; either a { Url, Description } object or a plain URL string. */
  LinkUrl: { Url: string; Description: string } | string;
  /** Optional manual description separate from the hyperlink Description field. */
  LinkDescription?: string;
  /** Optional category string used to filter results server-side. */
  Category?: string;
  /** Numeric sort order; results are ordered by this field ascending. */
  SortOrder?: number;
  /** ISO 8601 date string; link is hidden before this date when set. */
  StartDate?: string;
  /** ISO 8601 date string; link is hidden after this date when set. */
  StopDate?: string;
}

/**
 * Normalised link item used internally by the component.
 * Mapped from the raw SharePoint REST response by the `getLinks` function.
 */
interface ILinkItem {
  /** SharePoint item ID. */
  id: number;
  /** Display title / anchor text. */
  title: string;
  /** Target URL. */
  url: string;
  /** Optional description shown beneath the anchor. */
  description?: string;
  /** Numeric sort order (lower = earlier in list). */
  sortOrder?: number;
  /** Category label for potential filtering/grouping. */
  category?: string;
  /** Date from which this link should be visible. */
  startDate?: Date;
  /** Date after which this link should be hidden. */
  stopDate?: Date;
}

/**
 * isLinkActive
 *
 * Determines whether a link item should be displayed based on its start/stop
 * date window compared to today's date (time component stripped).
 *
 * @param item - Normalised link item to evaluate.
 * @returns    True when the item is within its visibility window.
 */
function isLinkActive(item: ILinkItem): boolean {
  const now = new Date();
  // Strip time component so comparisons are day-level only.
  now.setHours(0, 0, 0, 0);

  if (item.startDate) {
    const start = new Date(item.startDate);
    start.setHours(0, 0, 0, 0);
    // Link has not started yet.
    if (now < start) return false;
  }

  if (item.stopDate) {
    const stop = new Date(item.stopDate);
    stop.setHours(0, 0, 0, 0);
    // Link's display period has ended.
    if (now > stop) return false;
  }

  return true;
}

/**
 * Links
 *
 * Main exported component rendered by LinksWebPart (spfx-links-webpart).
 * Fetches links from the configured SharePoint list and renders them as an
 * unordered list, applying date-window filtering and optional category filter.
 *
 * @param props - Configuration and context forwarded from LinksWebPart.ts.
 */
export default function Links(props: ILinksProps): JSX.Element {
  /** Array of normalised ILinkItem objects currently displayed. */
  const [items, setItems] = useState<ILinkItem[]>([]);

  /** True while the data-fetch request is in-flight. */
  const [loading, setLoading] = useState<boolean>(false);

  /** Non-empty when a fetch error occurs; shown in red text above the list. */
  const [error, setError] = useState<string>('');

  /**
   * Data-fetch effect.
   * Re-runs when listId, maxItems, or category props change.
   * Guards against running when no list has been selected (avoids a 400 REST error).
   */
  useEffect(() => {
    // Skip fetch entirely when no list GUID is configured yet.
    if (!props.listId) {
      setItems([]);
      return;
    }

    const load = async (): Promise<void> => {
      setLoading(true);
      setError('');

      try {
        // Fetch and normalise items from the SharePoint REST endpoint.
        const data = await getLinks(
          props.spHttpClient,
          props.webUrl,
          props.listId,
          props.maxItems,
          props.category
        );
        setItems(data);
      } catch (e: unknown) {
        // Surface a readable error message; clear items to avoid showing stale data.
        const errorMessage = e instanceof Error ? e.message : 'Failed to load links.';
        setError(errorMessage);
        setItems([]);
      } finally {
        setLoading(false);
      }
    };

    load().catch(() => {
      // Error already handled in catch block above.
    });
  }, [props.listId, props.maxItems, props.category]);

  return (
    <div>
      {props.title && <h2>{props.title}</h2>}

      {!props.listId && <div>Please edit the web part and select a list.</div>}

      {loading && <div>Loading links...</div>}
      {error && <div style={{ color: 'red' }}>{error}</div>}

      {!loading && !error && items.length === 0 && props.listId && (
        <div>No links found.</div>
      )}

      <ul>
        {items.filter(isLinkActive).map(i => (
          <li key={i.id} style={{ marginBottom: '8px' }}>
            <a
              href={i.url}
              target={props.openInNewTab ? '_blank' : '_self'}
              rel={props.openInNewTab ? 'noreferrer noopener' : undefined}
            >
              {i.title}
            </a>

            {props.showDescription && i.description && (
              <div style={{ fontSize: '12px', opacity: 0.8 }}>{i.description}</div>
            )}
          </li>
        ))}
      </ul>
    </div>
  );
}

/**
 * getLinks
 *
 * Module-private async function that fetches and normalises link items from a
 * SharePoint list using the SPHttpClient REST API.
 *
 * The function:
 *  1. Builds an OData endpoint URL targeting the list by GUID.
 *  2. Optionally appends a $filter clause when a category is specified.
 *  3. Orders results by SortOrder ascending, then Title ascending.
 *  4. Maps the raw REST rows to typed ILinkItem objects.
 *  5. Filters out any items that have no URL (avoids rendering broken anchors).
 *
 * IMPORTANT: If your list uses different internal column names, update the
 * `select` string and the mapping in the `.map()` call accordingly.
 *
 * @param spHttpClient - Authenticated SPHttpClient from the SPFx context.
 * @param webUrl       - Absolute URL of the SharePoint site.
 * @param listId       - GUID of the target SharePoint list.
 * @param maxItems     - Maximum number of items to retrieve ($top value).
 * @param category     - Optional category filter string; empty = no filter.
 * @returns            Normalised array of active ILinkItem objects (URL present).
 */
async function getLinks(
  spHttpClient: SPHttpClient,
  webUrl: string,
  listId: string,
  maxItems: number,
  category: string
): Promise<ILinkItem[]> {

  // Select only the columns required for rendering to keep the payload small.
  // Replace column names here if your list uses different internal names.
  const select = `Id,Title,LinkUrl,LinkDescription,Category,SortOrder,StartDate,StopDate`;

  // Build an OData $filter clause only when a category filter has been configured.
  // escapeODataString prevents OData injection from single-quoted values.
  const filter = category
    ? `&$filter=Category eq '${escapeODataString(category)}'`
    : '';

  // Construct the full REST endpoint.
  // Lists are addressed by GUID to avoid breakage on list renames.
  const endpoint =
    `${webUrl}/_api/web/lists(guid'${listId}')/items` +
    `?$select=${select}` +
    `&$orderby=SortOrder asc,Title asc` +
    `&$top=${maxItems}` +
    filter;

  // Issue the GET request through SPHttpClient (credentials injected by SPFx runtime).
  // OData metadata=none reduces payload size by omitting type annotations.
  const res: SPHttpClientResponse = await spHttpClient.get(
    endpoint,
    SPHttpClient.configurations.v1,
    {
      headers: {
        'accept': 'application/json;odata.metadata=none'
      }
    }
  );

  if (!res.ok) {
    throw new Error(`SharePoint REST error: ${res.status} ${res.statusText}`);
  }

  const json = await res.json();
  const raw = json.value ?? [];

  return raw.map((r: ISharePointLinkItem) => {
    // Hyperlink columns can return either { Url, Description } or a plain string.
    const linkVal = r.LinkUrl;
    const url = typeof linkVal === 'string' ? linkVal : (linkVal?.Url ?? '');

    // Prefer the explicit LinkDescription field; fall back to the hyperlink column description.
    const description = r.LinkDescription ?? (typeof linkVal === 'object' ? linkVal?.Description : '') ?? '';

    // Parse dates if present (SharePoint returns ISO 8601 format).
    const startDate = r.StartDate ? new Date(r.StartDate) : undefined;
    const stopDate = r.StopDate ? new Date(r.StopDate) : undefined;

    return {
      id: r.Id,
      title: r.Title,
      url,
      description,
      category: r.Category,
      sortOrder: r.SortOrder,
      startDate,
      stopDate
    } as ILinkItem;
  // Remove items that have no URL; a link without a target cannot be rendered usefully.
  }).filter((x: ILinkItem) => !!x.url);
}

/**
 * escapeODataString
 *
 * Escapes single-quote characters in a string so it can be safely embedded
 * inside a single-quoted OData string literal (e.g. `$filter=Category eq '...'`).
 *
 * In OData, a single quote inside a string literal is escaped by doubling it.
 * Without this escaping, user-supplied values could break the filter syntax.
 *
 * @param input - The raw string to escape.
 * @returns     The input with every `'` replaced by `''`.
 */
function escapeODataString(input: string): string {
  return input.replace(/'/g, "''");
}
