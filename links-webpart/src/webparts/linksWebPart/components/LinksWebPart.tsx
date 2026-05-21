/**
 * LinksWebPart.tsx  (links-webpart)
 *
 * React functional component that renders a categorised, collapsible list of
 * hyperlinks sourced from a SharePoint list.  Each category is displayed as
 * an expandable <details> block; only one category can be open at a time.
 *
 * Features:
 *  - Loads link items via the PnP SP service (getLinks) on mount and when
 *    key props change (listTitle, maxItems).
 *  - Groups items by Category using getCategoryLabel to normalise the column
 *    value regardless of whether it is a plain string, managed-metadata term,
 *    array, or null.
 *  - Exposes configurable colours via CSS custom properties driven by props.
 *  - Optionally shows an IT Hotline callout box above the link list.
 *  - Shows item Description beneath each link anchor when showDescription=true.
 */
import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from './LinksWebPart.module.scss';
import type { ILinksWebPartProps } from './ILinksWebPartProps';
import { getLinks } from './services/linkService';
import type { ILinkItem } from './services/ILinkItem';

/** Fallback category label used when a link item has no Category value. */
const UNCATEGORIZED_LABEL = 'Uncategorized';

/**
 * Augmented CSSProperties type that includes the CSS custom properties used
 * by LinksWebPart.module.scss for runtime colour theming.
 * Declaring them here allows TypeScript to accept these properties on inline
 * `style` objects without casting to `any`.
 */
type LinksStyleVars = React.CSSProperties & {
  '--links-background': string;
  '--links-titlebar-bg': string;
  '--links-text-color': string;
  '--links-text-hover-color': string;
  '--links-text-visited-color': string;
};

/**
 * getCategoryLabel
 *
 * Normalises the raw Category column value from a SharePoint list item into
 * a plain display string.
 *
 * SharePoint can return a Category value in several formats depending on the
 * column type and API endpoint:
 *  - A plain string (simple text column).
 *  - An array (multi-value text or multi-value managed metadata).
 *  - An object with a `Label`, `TermLabel`, or `value` property (managed-metadata term).
 *  - null / undefined (field absent or empty).
 *
 * @param value - Raw column value from the SharePoint REST response.
 * @returns      Human-readable category label, or UNCATEGORIZED_LABEL.
 */
function getCategoryLabel(value: unknown): string {
  if (typeof value === 'string') {
    const trimmed = value.trim();
    // Non-empty string: use as-is.
    return trimmed || UNCATEGORIZED_LABEL;
  }

  if (Array.isArray(value)) {
    // Multi-value: recursively resolve each entry, then join non-Uncategorized labels.
    const labels = value
      .map((entry) => getCategoryLabel(entry))
      .filter((label) => label && label !== UNCATEGORIZED_LABEL);

    return labels.length ? labels.join(', ') : UNCATEGORIZED_LABEL;
  }

  if (value && typeof value === 'object') {
    // Managed-metadata term object: try known label property names in priority order.
    const candidate = value as { Label?: unknown; TermLabel?: unknown; value?: unknown };
    const rawLabel = candidate.Label ?? candidate.TermLabel ?? candidate.value;
    if (typeof rawLabel === 'string') {
      const trimmed = rawLabel.trim();
      if (trimmed) {
        return trimmed;
      }
    }
  }

  // Null, undefined, number, boolean, or unrecognised object — fall back.
  return UNCATEGORIZED_LABEL;
}

/**
 * LinksWebPart
 *
 * Main exported component.  Fetches link data from SharePoint, groups the
 * results by category, and renders them as collapsible sections.
 *
 * @param props - Configuration and context forwarded from LinksWebPartWebPart.
 */
export default function LinksWebPart(props: ILinksWebPartProps): JSX.Element {
  /** Array of ILinkItem objects fetched from the SharePoint list. */
  const [items, setItems] = useState<ILinkItem[]>([]);

  /** True while the initial data fetch is in-flight; controls loading indicator. */
  const [loading, setLoading] = useState<boolean>(true);

  /** Non-empty when a fetch error occurs; displayed as an error card instead of the list. */
  const [error, setError] = useState<string>('');

  /**
   * The currently expanded category label, or null when all sections are collapsed.
   * Only one category may be open at a time; clicking an open category closes it.
   */
  const [expandedCategory, setExpandedCategory] = useState<string | null>(null);

  /**
   * Data-fetch effect.
   * Runs after the initial render and whenever listTitle or maxItems change.
   * Uses an `isMounted` flag to guard against setting state after unmount
   * (prevents React's "can't perform a state update on an unmounted component" warning).
   */
  useEffect(() => {
    let isMounted = true;

    async function load(): Promise<void> {
      try {
        setLoading(true);
        setError('');

        // Fetch items from the configured SharePoint list via PnP SP.
        const data = await getLinks(props.context, props.listTitle, props.maxItems);

        if (isMounted) {
          setItems(data || []);
        }
      } catch (e: unknown) {
        if (isMounted) {
          // Extract a human-readable message; fall back to a generic string.
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

    // Cleanup: prevent state updates if the component unmounts before fetch completes.
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

  /**
   * Group items by their resolved category label.
   * The accumulator builds a dictionary keyed by category string; each value
   * is an array of ILinkItem objects belonging to that category.
   * Items with a null/empty Category fall into the UNCATEGORIZED_LABEL bucket.
   */
  const groupedItems = items.reduce<Record<string, ILinkItem[]>>((groups, item) => {
    const category = getCategoryLabel(item.Category);

    if (!groups[category]) {
      groups[category] = [];
    }

    groups[category].push(item);
    return groups;
  }, {});

  /**
   * Sorted category keys for deterministic rendering order.
   * Case-insensitive locale comparison ensures consistent ordering across
   * different user locales (e.g. 'Admin' and 'admin' sort adjacently).
   */
  const categories = Object.keys(groupedItems).sort((a, b) =>
    a.localeCompare(b, undefined, { sensitivity: 'base' })
  );

  /**
   * Inline style object carrying CSS custom property values.
   * The SCSS module reads these at runtime to apply the configured colours,
   * allowing property-pane colour pickers to change the theme without a
   * full re-bundle.
   */
  const containerStyle: LinksStyleVars = {
    '--links-background': props.backgroundColor,
    '--links-titlebar-bg': props.titleBarColor,
    '--links-text-color': props.linkTextColor,
    '--links-text-hover-color': props.linkHoverColor,
    '--links-text-visited-color': props.linkVisitedColor
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
