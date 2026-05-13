import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from './LinksWebPart.module.scss';
import type { ILinksWebPartProps } from './ILinksWebPartProps';
import { getLinks } from './services/linkService';
import type { ILinkItem } from './services/ILinkItem';

const UNCATEGORIZED_LABEL = 'Uncategorized';

type LinksStyleVars = React.CSSProperties & {
  '--links-background': string;
  '--links-titlebar-bg': string;
  '--links-text-color': string;
};

function getCategoryLabel(value: unknown): string {
  if (typeof value === 'string') {
    const trimmed = value.trim();
    return trimmed || UNCATEGORIZED_LABEL;
  }

  if (Array.isArray(value)) {
    const labels = value
      .map((entry) => getCategoryLabel(entry))
      .filter((label) => label && label !== UNCATEGORIZED_LABEL);

    return labels.length ? labels.join(', ') : UNCATEGORIZED_LABEL;
  }

  if (value && typeof value === 'object') {
    const candidate = value as { Label?: unknown; TermLabel?: unknown; value?: unknown };
    const rawLabel = candidate.Label ?? candidate.TermLabel ?? candidate.value;
    if (typeof rawLabel === 'string') {
      const trimmed = rawLabel.trim();
      if (trimmed) {
        return trimmed;
      }
    }
  }

  return UNCATEGORIZED_LABEL;
}

export default function LinksWebPart(props: ILinksWebPartProps): JSX.Element {
  const [items, setItems] = useState<ILinkItem[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>('');
  const [expandedCategory, setExpandedCategory] = useState<string | null>(null);

  useEffect(() => {
    let isMounted = true;

    async function load(): Promise<void> {
      try {
        setLoading(true);
        setError('');

        const data = await getLinks(props.context, props.listTitle, props.maxItems);

        if (isMounted) {
          setItems(data || []);
        }
      } catch (e: unknown) {
        if (isMounted) {
          const message = e instanceof Error ? e.message : 'Failed to load links.';
          setError(message);
        }
      } finally {
        if (isMounted) {
          setLoading(false);
        }
      }
    }

    load().catch(() => {
      // Error is already captured into component state in load().
    });

    return () => {
      isMounted = false;
    };
  }, [props.context, props.listTitle, props.maxItems]);

  if (loading) {
    return <div className={styles.status}>Loading links...</div>;
  }

  if (error) {
    return (
      <div className={styles.statusError}>
        <div className={styles.errorTitle}>Could not load links</div>
        <div className={styles.errorDetail}>{error}</div>
      </div>
    );
  }

  if (!items.length) {
    return <div className={styles.status}>No links found.</div>;
  }

  const groupedItems = items.reduce<Record<string, ILinkItem[]>>((groups, item) => {
    const category = getCategoryLabel(item.Category);

    if (!groups[category]) {
      groups[category] = [];
    }

    groups[category].push(item);
    return groups;
  }, {});

  const categories = Object.keys(groupedItems).sort((a, b) =>
    a.localeCompare(b, undefined, { sensitivity: 'base' })
  );

  const containerStyle: LinksStyleVars = {
    '--links-background': props.backgroundColor,
    '--links-titlebar-bg': props.titleBarColor,
    '--links-text-color': props.linkTextColor
  };

  return (
    <div className={styles.categoryList} style={containerStyle}>
      {props.showTopText && (
        <div className={styles.hotlineBox}>
          <div>If you are unable to enter an IT Help Desk Ticket, please call the IT Hotline</div>
          <div className={styles.hotlineLine}>
            <span>CCF &amp; ASB:</span>
            <span>240-223-3333</span>
          </div>
          <div className={styles.hotlineLine}>
            <span>BFS:</span>
            <span>301-986-6010</span>
          </div>
        </div>
      )}

      {categories.map((category) => (
        <details key={category} className={styles.categorySection} open={expandedCategory === category}>
          <summary
            className={styles.categorySummary}
            onClick={(event) => {
              event.preventDefault();
              setExpandedCategory((current) => (current === category ? null : category));
            }}
          >
            <span className={styles.categoryTitle}>{category}</span>
          </summary>
          <ul className={styles.linkList}>
            {groupedItems[category].map((item) => {
              const url = item.Link?.Url;
              const label = item.Title;

              if (!url) {
                return (
                  <li key={item.Id} className={styles.linkRow}>
                    <span className={styles.linkItemDisabled}>{item.Title}</span>
                    {props.showDescription && item.Description && (
                      <div className={styles.desc}>{item.Description}</div>
                    )}
                  </li>
                );
              }

              return (
                <li key={item.Id} className={styles.linkRow}>
                  <a
                    className={styles.linkItem}
                    href={url}
                    target={props.openInNewTab ? '_blank' : '_self'}
                    rel={props.openInNewTab ? 'noreferrer noopener' : undefined}
                    aria-label={label}
                  >
                    {label}
                  </a>
                  {props.showDescription && item.Description && (
                    <div className={styles.desc}>{item.Description}</div>
                  )}
                </li>
              );
            })}
          </ul>
        </details>
      ))}
    </div>
  );
}
