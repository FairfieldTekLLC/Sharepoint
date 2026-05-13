import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface ISiteHit {
  id: string;
  webUrl?: string;
  displayName?: string;
  name?: string;
  description?: string;
}

export default class SiteSearchService {
  constructor(private graphClient: MSGraphClientV3) {}

  public async searchSites(queryString: string, from = 0, size = 25): Promise<{ items: ISiteHit[]; more?: boolean; total?: number; }> {
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

  public async listAccessibleSites(max = 300): Promise<ISiteHit[]> {
    const pageSize = 100;
    const collected: ISiteHit[] = [];
    let from = 0;
    let more = true;

    while (more && collected.length < max) {
      const page = await this.searchSites('*', from, pageSize);
      const items = page.items ?? [];

      collected.push(...items);
      from += pageSize;
      more = Boolean(page.more) && items.length > 0;
    }

    const deduped = new Map<string, ISiteHit>();
    for (const site of collected) {
      const key = (site.webUrl || site.id || '').toLowerCase();
      if (!key || deduped.has(key)) {
        continue;
      }

      deduped.set(key, site);
    }

    return Array.from(deduped.values());
  }
}