import * as React from 'react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import SiteSearchService, { ISiteHit } from '../../services/SiteSearchService';

export default function CustomNavBar(props: { graphClient?: MSGraphClientV3 }): React.ReactElement {
  const service = React.useMemo(
    () => props.graphClient ? new SiteSearchService(props.graphClient) : undefined,
    [props.graphClient]
  );

  const [query, setQuery] = React.useState('');
  const [items, setItems] = React.useState<ISiteHit[]>([]);
  const [from, setFrom] = React.useState(0);
  const [more, setMore] = React.useState(false);
  const [loading, setLoading] = React.useState(false);
  const [searchUnavailable, setSearchUnavailable] = React.useState(false);

  const size = 25;

  const load = React.useCallback(async (reset: boolean) => {
    if (!service) {
      setSearchUnavailable(true);
      setMore(false);
      setItems([]);
      return;
    }

    setLoading(true);
    try {
      const nextFrom = reset ? 0 : from;
      const r = await service.searchSites(query, nextFrom, size);
      setSearchUnavailable(false);
      setMore(!!r.more);
      setFrom(nextFrom + size);
      setItems(prev => reset ? r.items : [...prev, ...r.items]);
    } catch {
      setSearchUnavailable(true);
      setMore(false);
      if (reset) {
        setItems([]);
      }
    } finally {
      setLoading(false);
    }
  }, [service, query, from]);

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
