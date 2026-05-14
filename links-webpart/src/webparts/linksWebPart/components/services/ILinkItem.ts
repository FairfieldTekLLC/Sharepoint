/**
 * ILinkItem.ts
 *
 * Data transfer object (DTO) types that model link items returned from the
 * SharePoint list REST API for the links-webpart project.
 */

/**
 * Represents the value of a SharePoint Hyperlink or Picture column.
 * SharePoint REST returns this column type as an object with `Url` and an
 * optional `Description` (the link's display text as entered by the list editor).
 */
export interface ILinkFieldValue {
  /** The target URL of the hyperlink. */
  Url: string;

  /**
   * Optional display text for the hyperlink as stored in the Hyperlink column.
   * If blank, the component falls back to the item's `Title` field.
   */
  Description?: string;
}

/**
 * Represents a single row from the SharePoint links list.
 * Field names must match the internal names of the corresponding list columns.
 */
export interface ILinkItem {
  /** SharePoint list item ID (auto-generated integer primary key). */
  Id: number;

  /** Display title of the link. Shown as the anchor text when Description is absent. */
  Title: string;

  /**
   * Hyperlink column value containing the target URL and optional description.
   * Mapped from the SharePoint 'Link' column (internal name).
   */
  Link: ILinkFieldValue;

  /**
   * Category column value used to group links into collapsible sections.
   * Typed as `unknown` because SharePoint can return a managed-metadata term,
   * a plain string, an array, or null depending on column configuration.
   * The `getCategoryLabel` helper in LinksWebPart.tsx normalises this to a string.
   */
  Category?: unknown;

  /** Optional human-readable description displayed beneath the link anchor. */
  Description?: string;

  /** Numeric sort order controlling the display sequence within a category. */
  Order?: number;

  /**
   * Boolean flag; only items where Active === 1 (truthy) are fetched.
   * Controlled via the `$filter=Active eq 1` OData clause in linkService.ts.
   */
  Active?: boolean;
}
