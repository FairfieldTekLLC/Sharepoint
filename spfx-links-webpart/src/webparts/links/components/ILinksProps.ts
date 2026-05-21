/**
 * ILinksProps.ts  (spfx-links-webpart)
 *
 * Defines the props contract for the Links React component.
 * All values are resolved and forwarded by the LinksWebPart SPFx web part class
 * from the property pane configuration and the active page context.
 */
import { SPHttpClient } from '@microsoft/sp-http';

/**
 * Props accepted by the Links functional component.
 *
 * Links renders a flat ordered list of hyperlinks fetched from a specific
 * SharePoint list identified by GUID.  Items outside their active date window
 * are filtered out client-side by `isLinkActive`.
 */
export interface ILinksProps {
  /** Display title rendered as a heading above the link list. */
  title: string;

  /**
   * Absolute URL of the SharePoint site (e.g. https://tenant.sharepoint.com/sites/Intranet).
   * Used as the base for all REST API calls inside the `getLinks` function.
   */
  webUrl: string;

  /**
   * The SPFx SPHttpClient instance.
   * Used to make authenticated SharePoint REST calls without requiring a separate
   * token acquisition step (the SPFx runtime injects credentials automatically).
   */
  spHttpClient: SPHttpClient;

  /**
   * GUID of the SharePoint list that contains the link items.
   * Using the GUID (rather than Title) avoids breakage if the list is renamed.
   * Populated from the property pane dropdown that enumerates visible lists.
   */
  listId: string;

  /**
   * Maximum number of list items to request from the REST endpoint.
   * Configurable from the property pane (range 1–50); defaults to 10.
   */
  maxItems: number;

  /**
   * When true, renders the `LinkDescription` field beneath each link anchor.
   * Toggle off for a compact, title-only list.
   */
  showDescription: boolean;

  /**
   * When true, each link opens in a new browser tab (`target="_blank"`).
   * When false, links navigate in the current tab (`target="_self"`).
   */
  openInNewTab: boolean;

  /**
   * Optional category filter string.
   * When non-empty, an OData `$filter=Category eq '...'` clause is appended
   * to the REST query so only items in the specified category are returned.
   * Leave blank to retrieve items from all categories.
   */
  category: string;

  /** Optional custom color applied to rendered link anchors. */
  linkColor?: string;

  /** Optional custom color applied to followed (visited) links. */
  followedLinkColor?: string;

  /** Optional secondary description text (currently unused by the Links component). */
  description?: string;

  /** True when SharePoint is using an inverted (dark) theme. Passed for future theming use. */
  isDarkTheme?: boolean;

  /** Human-readable environment label (e.g. 'Local', 'Remote'). For diagnostics / display. */
  environmentMessage?: string;

  /** True when the web part is running inside a Microsoft Teams context. */
  hasTeamsContext?: boolean;

  /** Display name of the currently signed-in user. Available for personalised greetings. */
  userDisplayName?: string;
}