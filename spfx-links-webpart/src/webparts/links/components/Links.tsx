import * as React from 'react';
import { useEffect, useState } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http'; // SPHttpClient 【5-d874ab】【6-7acd82】
import { ILinksProps } from './ILinksProps';

// You can use Fluent/Office UI components if your project includes them.
// If not, you can render plain <a> tags just fine.

interface ISharePointLinkItem {
  Id: number;
  Title: string;
  LinkUrl: { Url: string; Description: string } | string;
  LinkDescription?: string;
  Category?: string;
  SortOrder?: number;
  StartDate?: string;
  StopDate?: string;
}

interface ILinkItem {
  id: number;
  title: string;
  url: string;
  description?: string;
  sortOrder?: number;
  category?: string;
  startDate?: Date;
  stopDate?: Date;
}

function isLinkActive(item: ILinkItem): boolean {
  const now = new Date();
  now.setHours(0, 0, 0, 0);

  if (item.startDate) {
    const start = new Date(item.startDate);
    start.setHours(0, 0, 0, 0);
    if (now < start) return false;
  }

  if (item.stopDate) {
    const stop = new Date(item.stopDate);
    stop.setHours(0, 0, 0, 0);
    if (now > stop) return false;
  }

  return true;
}

export default function Links(props: ILinksProps): JSX.Element {
  const [items, setItems] = useState<ILinkItem[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>('');

  useEffect(() => {
    if (!props.listId) {
      setItems([]);
      return;
    }

    const load = async (): Promise<void> => {
      setLoading(true);
      setError('');

      try {
        const data = await getLinks(
          props.spHttpClient,
          props.webUrl,
          props.listId,
          props.maxItems,
          props.category
        );
        setItems(data);
      } catch (e: unknown) {
        const errorMessage = e instanceof Error ? e.message : 'Failed to load links.';
        setError(errorMessage);
        setItems([]);
      } finally {
        setLoading(false);
      }
    };

    load().catch(() => {
      // Error already handled in catch block
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

async function getLinks(
  spHttpClient: SPHttpClient,
  webUrl: string,
  listId: string,
  maxItems: number,
  category: string
): Promise<ILinkItem[]> {

  // Replace these field names with YOUR internal names if different:
  // Title, LinkUrl, LinkDescription, Category, SortOrder, StartDate, StopDate
  const select = `Id,Title,LinkUrl,LinkDescription,Category,SortOrder,StartDate,StopDate`;

  const filter = category
    ? `&$filter=Category eq '${escapeODataString(category)}'`
    : '';

  // List items REST by GUID 【7-1f2a7b】
  const endpoint =
    `${webUrl}/_api/web/lists(guid'${listId}')/items` +
    `?$select=${select}` +
    `&$orderby=SortOrder asc,Title asc` +
    `&$top=${maxItems}` +
    filter;

  // Use SPHttpClient to call SharePoint REST 【5-d874ab】【6-7acd82】
  const res: SPHttpClientResponse = await spHttpClient.get(
    endpoint,
    SPHttpClient.configurations.v1,
    {
      headers: {
        // Microsoft example for odata.metadata=none header 【6-7acd82】
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
    // Hyperlink columns often return { Url, Description }
    const linkVal = r.LinkUrl;
    const url = typeof linkVal === 'string' ? linkVal : (linkVal?.Url ?? '');

    const description = r.LinkDescription ?? (typeof linkVal === 'object' ? linkVal?.Description : '') ?? '';

    // Parse dates if present (SharePoint returns ISO 8601 format)
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
  }).filter((x: ILinkItem) => !!x.url);
}

function escapeODataString(input: string): string {
  return input.replace(/'/g, "''");
}
