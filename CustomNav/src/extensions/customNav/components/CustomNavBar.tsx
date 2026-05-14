/**
 * CustomNavBar.tsx
 *
 * React functional component that renders a lightweight navigation bar with
 * a live site search feature powered by the Microsoft Graph Search API.
 *
 * The component is instantiated by CustomNavApplicationCustomizer and receives
 * an optional MSGraphClientV3 instance.  When the client is unavailable (e.g.
 * Graph permissions not granted), it gracefully degrades to a static banner
 * with an explanatory message rather than crashing.
 *
 * Layout: a horizontal flex bar containing a title, a search input, a Search
 * button, a "More" pagination button (when applicable), and up to 8 result
 * links shown as a compact inline preview.
 */
import * as React from 'react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import SiteSearchService, { ISiteHit } from '../../services/SiteSearchService';

/**
 * CustomNavBar
 *
 * @param props.graphClient - Optional Graph v3 client.  Absence indicates that
 *                            the extension could not initialise Graph; the
 *                            component shows a degraded-mode message instead.
 */
export default function CustomNavBar(props: { graphClient?: MSGraphClientV3 }): React.ReactElement {
  /**
   * Memoised SiteSearchService instance.
   * Re-created only when graphClient changes (practically once per mount).
   * Returns undefined when graphClient is absent so callers can guard on it.
   */
  const service = React.useMemo(
    () => props.graphClient ? new SiteSearchService(props.graphClient) : undefined,
    [props.graphClient]
  );

  /** Current value of the search text input. */
  const [query, setQuery] = React.useState('');

  /** Array of site result objects currently displayed in the preview list. */
  const [items, setItems] = React.useState<ISiteHit[]>([]);

  /** Pagination offset tracking how many items have been loaded so far. */
  const [from, setFrom] = React.useState(0);

  /** True when additional result pages are available from the Graph Search API. */
  const [more, setMore] = React.useState(false);

  /** True while an async search request is in-flight; disables buttons to prevent double-submit. */
  const [loading, setLoading] = React.useState(false);

  /** True when Graph Search is unavailable (no client or API error); activates degraded-mode UI. */
  const [searchUnavailable, setSearchUnavailable] = React.useState(false);

  /** Number of results to request per page from SiteSearchService. */
  const size = 25;

  /**
   * load
   *
   * Executes a site search and updates component state with the results.
   *
   * @param reset - When true, clears existing results and resets the pagination
   *                offset to 0 before fetching (used for a new search query).
   *                When false, appends the next page to the existing list.
   */
  const load = React.useCallback(async (reset: boolean) => {
    // Service is absent when the Graph client was unavailable at mount time.
    if (!service) {
      setSearchUnavailable(true);
      setMore(false);
      setItems([]);
      return;
    }

    setLoading(true);
    try {
      // Use offset 0 for a fresh search; preserve `from` for "load more" paging.
      const nextFrom = reset ? 0 : from;
      const r = await service.searchSites(query, nextFrom, size);
      setSearchUnavailable(false);
      setMore(!!r.more);
      // Advance the pagination cursor by the page size for the next call.
      setFrom(nextFrom + size);
      // Replace items on reset; append on "load more".
      setItems(prev => reset ? r.items : [...prev, ...r.items]);
    } catch {
      // Any Graph error (network, permission) triggers degraded-mode UI.
      setSearchUnavailable(true);
      setMore(false);
      if (reset) {
        setItems([]);
      }
    } finally {
      setLoading(false);
    }
  }, [service, query, from]);

  /**
   * Initial load effect.
   * Triggers a wildcard search on first mount so the bar immediately shows
   * available sites without requiring the user to type anything.
   * The empty dependency array means this runs once after the initial render.
   */
  React.useEffect(() => {
    load(true).catch(() => undefined);
  }, []);

  return (
    <div style={{ display:'flex', alignItems:'center', gap:12, padding:'8px 12px', background:'#0b1220', color:'#fff' }}>
      <strong>CustomNav</strong>

      {searchUnavailable ? (
        <div style={{ marginLeft: 'auto', fontSize: 13, opacity: 0.9 }}>
          Search is temporarily unavailable. Basic navigation is still active.
        </div>
      ) : (
        <>
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Search sites you can access"
            style={{ width: 320, padding:'6px 8px', borderRadius:8 }}
          />

          <button onClick={() => { setFrom(0); load(true).catch(() => undefined); }} disabled={loading}>
            Search
          </button>

          <div style={{ marginLeft:'auto', display:'flex', gap:8 }}>
            {more && <button onClick={() => { load(false).catch(() => undefined); }} disabled={loading}>More</button>}
          </div>

          {/* Small preview list – for real UX use a panel/dropdown */}
          <div style={{ display:'flex', gap:8, flexWrap:'wrap', maxWidth:700 }}>
            {items.slice(0, 8).map(s => (
              <a key={s.id} href={s.webUrl} style={{ color:'#93c5fd' }}>
                {s.displayName || s.name || s.webUrl}
              </a>
            ))}
          </div>
        </>
      )}
    </div>
  );
}
