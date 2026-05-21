/**
 * ILinksWebPartProps.ts
 *
 * Defines the props contract for the LinksWebPart React component (links-webpart).
 * All values are supplied by the SPFx web part class (LinksWebPartWebPart) from the
 * property pane configuration and the active WebPartContext.
 */
import { WebPartContext } from "@microsoft/sp-webpart-base";

/**
 * Props accepted by the LinksWebPart functional component.
 *
 * LinksWebPart renders a categorised, collapsible list of hyperlinks
 * sourced from a SharePoint list.  Appearance is fully customisable via
 * CSS custom properties driven by the color props below.
 */
export interface ILinksWebPartProps {
  /**
   * The SPFx WebPartContext for the current page.
   * Used by the PnP SP helper (spfi/SPFx) to construct an authenticated
   * SharePoint REST client scoped to the current site.
   */
  context: WebPartContext;

  /**
   * Display name of the SharePoint list that contains the link items.
   * Must match the list's Title exactly.
   * Configurable from the property pane; defaults to "Links".
   */
  listTitle: string;

  /**
   * Maximum number of link items to retrieve from the SharePoint list.
   * Configurable from the property pane (range 1–200); defaults to 12.
   */
  maxItems: number;

  /**
   * When true, each link opens in a new browser tab (`target="_blank"`).
   * When false, links navigate in the current tab (`target="_self"`).
   * Configurable from the property pane; defaults to true.
   */
  openInNewTab: boolean;

  /**
   * When true, renders the IT Hotline information box above the link list.
   * Toggle off to hide the static hotline text block.
   * Configurable from the property pane; defaults to true.
   */
  showTopText: boolean;

  /**
   * When true, shows the `Description` field beneath each link anchor.
   * Toggle off for a compact, title-only list.
   * Configurable from the property pane; defaults to true.
   */
  showDescription: boolean;

  /**
   * CSS color value for the overall card background.
   * Applied as the `--links-background` CSS custom property.
   * Configurable via color picker in the property pane; defaults to '#f7f7f7'.
   */
  backgroundColor: string;

  /**
   * CSS color value for the category header (title bar) background.
   * Applied as the `--links-titlebar-bg` CSS custom property.
   * Configurable via color picker in the property pane; defaults to '#d1d3aa'.
   */
  titleBarColor: string;

  /**
   * CSS color value for anchor text inside the link list.
   * Applied as the `--links-text-color` CSS custom property.
   * Configurable via color picker in the property pane; defaults to '#0e0e0e'.
   */
  linkTextColor: string;

  /**
   * CSS color value for anchor text on hover inside the link list.
   * Applied as the `--links-text-hover-color` CSS custom property.
   * Configurable via color picker in the property pane; defaults to '#0645ad'.
   */
  linkHoverColor: string;

  /**
   * CSS color value for visited anchor text inside the link list.
   * Applied as the `--links-text-visited-color` CSS custom property.
   * Configurable via color picker in the property pane; defaults to '#551a8b'.
   */
  linkVisitedColor: string;
}