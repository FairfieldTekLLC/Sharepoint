/**
 * CustomNavApplicationCustomizer.ts
 *
 * SPFx Application Customizer that injects a tenant-aware navigation bar into
 * the current SharePoint page. The extension prefers the Top placeholder,
 * falls back to Bottom when Top is unavailable, and finally falls back to a
 * fixed body-level host so the navigation still appears on pages without a
 * compatible placeholder surface.
 *
 * High-level responsibilities:
 * - Acquire a Microsoft Graph client and load accessible SharePoint sites.
 * - Cache the site list in localStorage for a short period to reduce Graph calls.
 * - Filter personal/team sites according to extension properties.
 * - Optionally render only configured custom nav items without loading sites.
 * - Build a tree from site URLs and render nested dropdown/flyout menus.
 * - Inject the CSS required for layout and interaction once per page.
 * - Bind click/keyboard handlers for menu open/close behavior.
 */
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import SiteSearchService, { ISiteHit } from '../services/SiteSearchService';

/**
 * Property bag for the CustomNav extension.
 *
 * The extension accepts a small set of optional properties for filtering the
 * built-in site list and configuring additional custom links.
 */
export interface ICustomNavApplicationCustomizerProperties {
  /** When true, skip Graph/site loading and render only custom nav items. */
  showOnlyCustomNavMenuItems?: boolean;
  /** Optional external links rendered alongside or instead of site-derived items. */
  externalLinks?: IExternalLinkConfig[];
  /** Hide OneDrive/personal sites when true. Defaults to true when omitted. */
  hidePersonalSites?: boolean;
  /** Hide team sites matching configured path prefixes when true. Defaults to true when omitted. */
  hideTeamSites?: boolean;
  /** Path prefixes that identify team sites, for example ['/teams/']. */
  teamSitePathPrefixes?: string[];
}

/** Node in the derived site tree used to render nested navigation menus. */
interface ISiteNode {
  /** Site represented by the current menu node. */
  site: ISiteHit;
  /** Immediate child sites whose URL path is nested under the current site. */
  children: ISiteNode[];
}

/** Configurable external link rendered alongside the site tree. */
interface IExternalLink {
  /** Link text displayed to the user. */
  title: string;
  /** Absolute http/https URL. */
  url: string;
  /** Browser target for navigation. */
  target: '_self' | '_blank';
  /** Optional nested static links rendered as dropdown/flyout children. */
  children: IExternalLink[];
}

/** Raw external-link config read from extension properties before validation/normalization. */
interface IExternalLinkConfig {
  title?: string;
  url?: string;
  target?: string;
  children?: IExternalLinkConfig[];
}

/** Cached site payload stored in localStorage. */
interface ICachedMenuData {
  /** Epoch milliseconds when the cache entry was written. */
  timestamp: number;
  /** Site results returned from Graph and filtered later at render time. */
  sites: ISiteHit[];
}

/** Window augmentation used to prevent duplicate extension initialisation on a page. */
interface ICustomNavWindow extends Window {
  __customNavInitialized?: boolean;
}

/**
 * CustomNavApplicationCustomizer
 *
 * Main extension class. Keeps enough DOM state to re-render when placeholders
 * change and when cached/remote site data becomes available.
 */
export default class CustomNavApplicationCustomizer
  extends BaseApplicationCustomizer<ICustomNavApplicationCustomizerProperties> {

  /** Cache lifetime for the locally stored site list: 2 minutes. */
  private static readonly MENU_CACHE_TTL_MS = 2 * 60 * 1000;

  /** Placeholder host provided by SPFx when Top or Bottom is available. */
  private _placeholder: PlaceholderContent | undefined;
  /** Fixed-position fallback host used when no SPFx placeholder is available. */
  private _fallbackHost: HTMLDivElement | undefined;
  /** Root element containing the navigation DOM tree. */
  private _navHost: HTMLDivElement | undefined;
  /** Injected <style> element containing the menu CSS. */
  private _styleTag: HTMLStyleElement | undefined;
  /** Prevents binding the same click/keyboard handlers more than once. */
  private _menuInteractionsBound = false;

  /**
   * Placeholder change callback.
   * Re-renders when SPFx reports that placeholder availability has changed.
   */
  private _onPlaceholdersChanged = (): void => {
    this._render().catch(() => undefined);
  };

  /**
   * onInit()
   *
   * SPFx lifecycle hook called once when the extension is initialised.
   * Guards against duplicate initialisation on the same page, subscribes to
   * placeholder changes, and triggers the first render.
   */
  public onInit(): Promise<void> {
    const navWindow = window as ICustomNavWindow;
    if (navWindow.__customNavInitialized) {
      console.warn('CustomNav: duplicate extension instance detected; skipping init.');
      return Promise.resolve();
    }

    navWindow.__customNavInitialized = true;
    console.log('CustomNav: onInit fired');
    this.context.placeholderProvider.changedEvent.add(this, this._onPlaceholdersChanged); // placeholder pattern 【1-551094】
    this._render().catch(() => undefined);
    return Promise.resolve();
  }

  /**
   * _render()
   *
   * Resolves the best available mount host, renders an immediate fallback shell,
   * then upgrades the UI once a Graph client becomes available.
   */
  private _render = async (): Promise<void> => {
    let mountHost: HTMLElement;

    if (!this._placeholder) {
      // Reset placeholder-specific DOM references when SPFx disposes the host.
      const onDispose = (): void => {
        this._navHost?.remove();
        this._navHost = undefined;
        this._placeholder = undefined;
      };

      this._placeholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose }
      );

      if (!this._placeholder) {
        this._placeholder = this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Bottom,
          { onDispose }
        );
      }

      if (!this._placeholder) {
        console.warn('CustomNav: no Top/Bottom placeholder found; using body fallback host.');
      }
    }

    if (this._placeholder) {
      // Prefer the framework placeholder host when available and remove any stale fallback host.
      if (this._fallbackHost) {
        this._fallbackHost.remove();
        this._fallbackHost = undefined;
      }
      mountHost = this._placeholder.domElement;
    } else {
      if (!this._fallbackHost) {
        // Body fallback keeps navigation visible even on pages without placeholders.
        this._fallbackHost = document.createElement('div');
        this._fallbackHost.id = 'custom-nav-fallback-host';
        this._fallbackHost.style.position = 'fixed';
        this._fallbackHost.style.top = '0';
        this._fallbackHost.style.left = '0';
        this._fallbackHost.style.right = '0';
        this._fallbackHost.style.zIndex = '100000';
        document.body.appendChild(this._fallbackHost);
      }

      mountHost = this._fallbackHost;
    }

    // Always render immediately; do not block nav visibility on Graph initialization.
    this._renderNav(mountHost, undefined);

    try {
      const graphClient: MSGraphClientV3 =
        await this.context.msGraphClientFactory.getClient('3');
      this._renderNav(mountHost, graphClient);
    } catch (error) {
      // Keep fallback nav visible even if Graph cannot be initialized.
      console.warn('CustomNav: Graph client unavailable, keeping fallback nav.', error);
    }
  };

  /**
   * _renderNav()
   *
   * Renders the navigation shell for one of three states:
   * - no Graph client yet: temporary "unavailable" placeholder
   * - cached data available: immediate menu render from local cache
   * - remote load required: loading state followed by async menu render
   */
  private _renderNav(container: HTMLElement, graphClient?: MSGraphClientV3): void {
    if (!this._navHost) {
      this._navHost = document.createElement('div');
      this._navHost.id = 'custom-nav-root';
    }

    this._ensureStyles();

    if (this._shouldShowOnlyCustomNavMenuItems()) {
      const externalLinks = this._getExternalLinks();

      this._navHost.innerHTML = '';
      if (!externalLinks.length) {
        const empty = document.createElement('div');
        empty.className = 'custom-nav-message';
        empty.textContent = 'No custom navigation items are configured.';
        this._navHost.appendChild(empty);
      } else {
        const list = document.createElement('ul');
        list.className = 'custom-nav-list custom-nav-list-root';
        externalLinks.forEach((link) => list.appendChild(this._renderExternalLinkItem(link, 0)));
        this._navHost.appendChild(list);
        this._bindMenuInteractions();
      }

      container.replaceChildren(this._navHost);
      return;
    }

    this._navHost.innerHTML = '';
    this._navHost.className = 'custom-nav-shell';

    if (!graphClient) {
      const unavailable = document.createElement('div');
      unavailable.className = 'custom-nav-message';
      unavailable.textContent = 'Navigation data is temporarily unavailable.';
      this._navHost.appendChild(unavailable);
      container.replaceChildren(this._navHost);
      return;
    }

    const service = new SiteSearchService(graphClient);

    // Prefer warm cache for faster first paint and fewer Graph requests.
    const cachedSites = this._getCachedSites();
    if (cachedSites) {
      this._renderSiteMenu(cachedSites);
      container.replaceChildren(this._navHost);
      return;
    }

    const loading = document.createElement('div');
    loading.className = 'custom-nav-message';
    loading.textContent = 'Loading sites...';
    this._navHost.appendChild(loading);
    container.replaceChildren(this._navHost);

    service.listAccessibleSites()
      .then((sites) => {
        this._setCachedSites(sites);
        this._renderSiteMenu(sites);
      })
      .catch((error) => {
        console.warn('CustomNav: failed to load sites for menu.', error);
        this._navHost!.innerHTML = '';
        const failed = document.createElement('div');
        failed.className = 'custom-nav-message';
        failed.textContent = 'Unable to load site navigation.';
        this._navHost!.appendChild(failed);
      });
  }

  /**
   * _renderSiteMenu()
   *
   * Normalises and filters the site set, builds the hierarchical tree, then
   * renders root site nodes and configured external links into the nav host.
   */
  private _renderSiteMenu(sites: ISiteHit[]): void {
    if (!this._navHost) {
      return;
    }

    const validSites = sites
      .filter((s) => Boolean(s.webUrl))
      .map((s) => ({ ...s, webUrl: this._normalizeUrl(s.webUrl!) }))
      .filter((s) => this._shouldIncludeSite(s.webUrl!));

    const tree = this._buildSiteTree(validSites);
    const externalLinks = this._getExternalLinks();

    this._navHost.innerHTML = '';

    const list = document.createElement('ul');
    list.className = 'custom-nav-list custom-nav-list-root';
    tree.forEach((node) => list.appendChild(this._renderSiteNode(node, 0)));
    externalLinks.forEach((link) => list.appendChild(this._renderExternalLinkItem(link, 0)));

    this._navHost.appendChild(list);
    this._bindMenuInteractions();
  }

  /**
   * _getExternalLinks()
   *
   * Reads the optional externalLinks property, trims values, normalises target,
   * and keeps only well-formed http/https links with non-empty titles.
   */
  private _getExternalLinks(): IExternalLink[] {
    const links = this.properties.externalLinks;

    if (!Array.isArray(links)) {
      return [];
    }

    return links
      .map((link) => this._normalizeExternalLink(link))
      .filter((link): link is IExternalLink => Boolean(link));
  }

  /** Validates and normalizes one external link entry, including optional nested children. */
  private _normalizeExternalLink(link: IExternalLinkConfig): IExternalLink | undefined {
    const target: '_self' | '_blank' = link.target === '_self' ? '_self' : '_blank';
    const normalized: IExternalLink = {
      title: (link.title || '').trim(),
      url: (link.url || '').trim(),
      target,
      children: Array.isArray(link.children)
        ? link.children
          .map((child) => this._normalizeExternalLink(child))
          .filter((child): child is IExternalLink => Boolean(child))
        : []
    };

    if (!normalized.title || !normalized.url) {
      return undefined;
    }

    if (!/^https?:\/\//i.test(normalized.url)) {
      return undefined;
    }

    return normalized;
  }

  /** Renders one external-link node recursively with dropdown/flyout support. */
  private _renderExternalLinkItem(link: IExternalLink, depth: number): HTMLLIElement {
    const item = document.createElement('li');
    item.className = 'custom-nav-item custom-nav-item-external';

    const hasChildren = link.children.length > 0;
    if (hasChildren) {
      item.classList.add('has-children');
    }

    const row = document.createElement('div');
    row.className = 'custom-nav-row';

    const anchor = document.createElement('a');
    anchor.className = depth === 0
      ? 'custom-nav-link custom-nav-link-external'
      : 'custom-nav-dropdown-link custom-nav-link-external';
    anchor.href = link.url;
    anchor.textContent = link.title;
    anchor.title = link.title;
    anchor.target = link.target;
    if (link.target === '_blank') {
      anchor.rel = 'noopener noreferrer';
    }

    row.appendChild(anchor);

    if (hasChildren) {
      const toggle = document.createElement('button');
      toggle.type = 'button';
      toggle.className = depth === 0 ? 'custom-nav-toggle' : 'custom-nav-flyout-toggle';
      toggle.setAttribute('aria-label', `Toggle submenu for ${link.title}`);
      toggle.setAttribute('aria-expanded', 'false');
      toggle.textContent = depth === 0 ? '▾' : '▸';
      row.appendChild(toggle);
    }

    item.appendChild(row);

    if (hasChildren) {
      const dropdown = document.createElement('ul');
      dropdown.className = 'custom-nav-dropdown';
      dropdown.setAttribute('data-depth', String(depth + 1));
      link.children.forEach((child) => dropdown.appendChild(this._renderExternalLinkItem(child, depth + 1)));
      item.appendChild(dropdown);
    }

    return item;
  }

  /**
   * _buildSiteTree()
   *
   * Converts a flat list of sites into a parent/child tree based on URL path
   * nesting. Parents are discovered with _findParentUrl(); nodes are then sorted
   * alphabetically by their display label at every depth.
   */
  private _buildSiteTree(sites: ISiteHit[]): ISiteNode[] {
    const byUrl = new Map<string, ISiteNode>();
    sites.forEach((site) => {
      byUrl.set(site.webUrl!, { site, children: [] });
    });

    const roots: ISiteNode[] = [];
    const knownUrls = new Set(byUrl.keys());

    byUrl.forEach((node, url) => {
      const parentUrl = this._findParentUrl(url, knownUrls);
      if (!parentUrl) {
        roots.push(node);
        return;
      }

      const parent = byUrl.get(parentUrl);
      if (!parent) {
        roots.push(node);
        return;
      }

      parent.children.push(node);
    });

    const sortNodes = (nodes: ISiteNode[]): void => {
      nodes.sort((a, b) => this._siteLabel(a.site).localeCompare(this._siteLabel(b.site)));
      nodes.forEach((n) => sortNodes(n.children));
    };

    sortNodes(roots);
    return roots;
  }

  /**
   * _renderSiteNode()
   *
   * Renders one site node recursively. Root nodes use the top-level link/toggle
   * classes; nested nodes use dropdown/flyout styles.
   */
  private _renderSiteNode(node: ISiteNode, depth: number): HTMLLIElement {
    const item = document.createElement('li');
    item.className = 'custom-nav-item';

    const hasChildren = node.children.length > 0;
    if (hasChildren) {
      item.classList.add('has-children');
    }

    const row = document.createElement('div');
    row.className = 'custom-nav-row';

    const link = document.createElement('a');
    link.className = depth === 0 ? 'custom-nav-link' : 'custom-nav-dropdown-link';
    link.href = node.site.webUrl || '#';
    link.textContent = this._siteLabel(node.site);
    link.title = this._siteLabel(node.site);
    row.appendChild(link);
    item.appendChild(row);

    if (hasChildren) {
      const toggle = document.createElement('button');
      toggle.type = 'button';
      toggle.className = depth === 0 ? 'custom-nav-toggle' : 'custom-nav-flyout-toggle';
      toggle.setAttribute('aria-label', `Toggle submenu for ${this._siteLabel(node.site)}`);
      toggle.setAttribute('aria-expanded', 'false');
      toggle.textContent = depth === 0 ? '▾' : '▸';
      row.appendChild(toggle);

      const dropdown = document.createElement('ul');
      dropdown.className = 'custom-nav-dropdown';
      dropdown.setAttribute('data-depth', String(depth + 1));
      node.children.forEach((child) => dropdown.appendChild(this._renderSiteNode(child, depth + 1)));
      item.appendChild(dropdown);
    }

    return item;
  }

  /**
   * _bindMenuInteractions()
   *
   * Attaches one-time event handlers that manage menu open/close behavior.
   * Clicks on toggles open a submenu, clicks outside close all menus, and Escape
   * also closes all open menus.
   */
  private _bindMenuInteractions(): void {
    if (!this._navHost || this._menuInteractionsBound) {
      return;
    }

    this._menuInteractionsBound = true;

    this._navHost.addEventListener('click', (event: MouseEvent) => {
      const target = event.target as HTMLElement;

      const link = target.closest('.custom-nav-link, .custom-nav-dropdown-link') as HTMLAnchorElement | null;
      if (link) {
        this._closeAllMenus();
        return;
      }

      const toggle = target.closest('.custom-nav-toggle, .custom-nav-flyout-toggle') as HTMLButtonElement | null;
      if (!toggle) {
        return;
      }

      event.preventDefault();
      event.stopPropagation();

      const item = toggle.closest('.custom-nav-item') as HTMLElement | null;
      if (!item) {
        return;
      }

      const expanded = item.classList.toggle('is-open');
      toggle.setAttribute('aria-expanded', String(expanded));

      if (expanded) {
        this._closeSiblingMenus(item);
      }
    });

    document.addEventListener('click', (event: MouseEvent) => {
      if (!this._navHost) {
        return;
      }

      const target = event.target as Node;
      if (!this._navHost.contains(target)) {
        this._closeAllMenus();
      }
    });

    document.addEventListener('keydown', (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        this._closeAllMenus();
      }
    });
  }

  /** Closes sibling menus when one menu item is opened. */
  private _closeSiblingMenus(item: HTMLElement): void {
    const parent = item.parentElement;
    if (!parent) {
      return;
    }

    Array.from(parent.children).forEach((child) => {
      if (child === item) {
        return;
      }

      const sibling = child as HTMLElement;
      sibling.classList.remove('is-open');
      sibling.querySelectorAll('.custom-nav-toggle, .custom-nav-flyout-toggle')
        .forEach((t) => t.setAttribute('aria-expanded', 'false'));
    });
  }

  /** Closes every open menu and resets aria-expanded on all toggles. */
  private _closeAllMenus(): void {
    if (!this._navHost) {
      return;
    }

    this._navHost.querySelectorAll('.custom-nav-item.is-open').forEach((item) => {
      item.classList.remove('is-open');
    });

    this._navHost.querySelectorAll('.custom-nav-toggle, .custom-nav-flyout-toggle').forEach((toggle) => {
      toggle.setAttribute('aria-expanded', 'false');
    });
  }

  /** Builds the localStorage cache key, scoped by site, user, and active filters. */
  private _getCacheKey(): string {
    const userKey = (this.context.pageContext.user.loginName || 'anonymous').toLowerCase();
    const siteKey = this.context.pageContext.site.absoluteUrl.toLowerCase();
    const filterKey = this._getFilterCacheKey();
    return `custom-nav-menu-cache::${siteKey}::${userKey}::${filterKey}`;
  }

  /** Encodes filter settings into the cache key so different filter combinations do not share stale data. */
  private _getFilterCacheKey(): string {
    const filterProps = this.properties;
    const hidePersonal = filterProps.hidePersonalSites !== false;
    const hideTeam = filterProps.hideTeamSites !== false;
    const prefixes = this._getTeamPathPrefixes().join(',');
    return `${hidePersonal ? 'hp1' : 'hp0'}-${hideTeam ? 'ht1' : 'ht0'}-${prefixes}`;
  }

  /** Returns true when the nav should render only custom links and skip site loading. */
  private _shouldShowOnlyCustomNavMenuItems(): boolean {
    return this.properties.showOnlyCustomNavMenuItems === true;
  }

  /** Applies personal-site and team-site filtering rules to a site URL. */
  private _shouldIncludeSite(url: string): boolean {
    const filterProps = this.properties;
    const hidePersonalSites = filterProps.hidePersonalSites !== false;
    const hideTeamSites = filterProps.hideTeamSites !== false;
    const parsed = new URL(url);
    const host = parsed.hostname.toLowerCase();
    const path = parsed.pathname.toLowerCase();

    if (hidePersonalSites) {
      const isPersonalHost = host.includes('-my.sharepoint.com');
      const isPersonalPath = path.startsWith('/personal/');
      if (isPersonalHost || isPersonalPath) {
        return false;
      }
    }

    if (hideTeamSites) {
      const teamPrefixes = this._getTeamPathPrefixes();
      if (teamPrefixes.some((prefix) => path.startsWith(prefix))) {
        return false;
      }
    }

    return true;
  }

  /** Normalises configured team-site path prefixes into lowercase leading-slash paths. */
  private _getTeamPathPrefixes(): string[] {
    const filterProps = this.properties;
    const configured = Array.isArray(filterProps.teamSitePathPrefixes)
      ? filterProps.teamSitePathPrefixes
      : ['/teams/'];

    return configured
      .map((prefix) => (prefix || '').trim().toLowerCase())
      .filter((prefix) => prefix.length > 0)
      .map((prefix) => (prefix.startsWith('/') ? prefix : `/${prefix}`));
  }

  /** Reads a non-expired cached site list from localStorage when available. */
  private _getCachedSites(): ISiteHit[] | undefined {
    try {
      const raw = window.localStorage.getItem(this._getCacheKey());
      if (!raw) {
        return undefined;
      }

      const parsed = JSON.parse(raw) as ICachedMenuData;
      if (!parsed?.timestamp || !Array.isArray(parsed.sites)) {
        return undefined;
      }

      const age = Date.now() - parsed.timestamp;
      if (age > CustomNavApplicationCustomizer.MENU_CACHE_TTL_MS) {
        return undefined;
      }

      return parsed.sites;
    } catch {
      return undefined;
    }
  }

  /** Writes the current site list to localStorage, ignoring storage failures silently. */
  private _setCachedSites(sites: ISiteHit[]): void {
    try {
      const payload: ICachedMenuData = {
        timestamp: Date.now(),
        sites
      };

      window.localStorage.setItem(this._getCacheKey(), JSON.stringify(payload));
    } catch {
      // Ignore storage errors to avoid impacting nav rendering.
    }
  }

  /** Resolves the best display label for a site. */
  private _siteLabel(site: ISiteHit): string {
    return site.displayName || site.name || site.webUrl || 'Site';
  }

  /** Normalises URLs for stable comparisons by trimming a trailing slash and lowercasing. */
  private _normalizeUrl(url: string): string {
    return url.replace(/\/$/, '').toLowerCase();
  }

  /**
   * Finds the nearest known parent URL by walking the current URL path upward.
   * Returns undefined when the site is a root node.
   */
  private _findParentUrl(url: string, knownUrls: Set<string>): string | undefined {
    const parsed = new URL(url);
    const parts = parsed.pathname.split('/').filter(Boolean);

    for (let i = parts.length - 1; i > 0; i--) {
      const candidatePath = '/' + parts.slice(0, i).join('/');
      const candidate = `${parsed.origin}${candidatePath}`.toLowerCase();
      if (knownUrls.has(candidate)) {
        return candidate;
      }
    }

    return undefined;
  }

  /**
   * _ensureStyles()
   *
   * Injects the CSS required by the navigation exactly once per page instance.
   * The stylesheet covers desktop dropdowns, nested flyouts, hover behavior, and
   * a mobile-friendly stacked layout for narrower viewports.
   */
  private _ensureStyles(): void {
    if (this._styleTag) {
      return;
    }

    this._styleTag = document.createElement('style');
    this._styleTag.id = 'custom-nav-style';
    this._styleTag.textContent = `
      .custom-nav-shell {
        background: #101a48;
        color: #ffffff;
        font-family: Segoe UI, Arial, sans-serif;
        border-bottom: 1px solid #0b1230;
        position: relative;
        z-index: 100001;
      }

      .custom-nav-list {
        list-style: none;
        margin: 0;
        padding: 8px 12px;
        display: flex;
        gap: 2px;
        flex-wrap: wrap;
        align-items: center;
        overflow: visible;
      }

      .custom-nav-item {
        position: relative;
        z-index: 1;
      }

      .custom-nav-item.is-open {
        z-index: 40;
      }

      @media (hover: hover) and (pointer: fine) {
        .custom-nav-item:hover {
          z-index: 40;
        }
      }

      .custom-nav-row {
        display: flex;
        align-items: center;
      }

      .custom-nav-link,
      .custom-nav-dropdown-link {
        color: #ffffff;
        text-decoration: none;
        display: block;
        white-space: nowrap;
      }

      .custom-nav-link {
        padding: 4px 8px;
        font-size: 14px;
        border-radius: 2px;
      }

      .custom-nav-link:hover {
        background: #1f2a63;
      }

      .custom-nav-link-external::after {
        content: ' ↗';
        font-size: 11px;
        opacity: 0.9;
      }

      .custom-nav-toggle,
      .custom-nav-flyout-toggle {
        border: 0;
        background: transparent;
        color: #ffffff;
        cursor: pointer;
        padding: 4px 4px;
        line-height: 1;
        transition: transform 0.18s ease;
      }

      .custom-nav-flyout-toggle {
        font-size: 11px;
        padding-right: 8px;
      }

      .custom-nav-item.is-open > .custom-nav-row > .custom-nav-toggle {
        transform: rotate(180deg);
      }

      .custom-nav-item.is-open > .custom-nav-row > .custom-nav-flyout-toggle {
        transform: rotate(90deg);
      }

      .custom-nav-dropdown {
        display: none;
        position: absolute;
        top: 100%;
        left: 0;
        min-width: 240px;
        margin: 0;
        padding: 6px 0;
        list-style: none;
        background: #1a2558;
        border: 1px solid #0b1230;
        box-shadow: 0 8px 20px rgba(0, 0, 0, 0.25);
        z-index: 50;
      }

      .custom-nav-dropdown .custom-nav-dropdown {
        top: -7px;
        left: 100%;
      }

      .custom-nav-item.is-open > .custom-nav-dropdown {
        display: block;
      }

      @media (hover: hover) and (pointer: fine) {
        .custom-nav-item:hover > .custom-nav-dropdown {
          display: block;
        }
      }

      .custom-nav-dropdown-link {
        padding: 5px 10px;
        font-size: 13px;
      }

      .custom-nav-dropdown-link:hover {
        background: #2b3a7f;
      }

      .custom-nav-message {
        padding: 10px 12px;
        font-size: 13px;
        opacity: 0.95;
      }

      @media (max-width: 900px) {
        .custom-nav-list-root {
          display: block;
          padding: 6px 0;
        }

        .custom-nav-item {
          width: 100%;
        }

        .custom-nav-link,
        .custom-nav-dropdown-link {
          flex: 1;
          padding: 10px 12px;
        }

        .custom-nav-dropdown,
        .custom-nav-dropdown .custom-nav-dropdown {
          position: static;
          min-width: 0;
          box-shadow: none;
          border-left: 2px solid #2b3a7f;
          margin-left: 10px;
        }

        .custom-nav-flyout-toggle {
          transform: rotate(90deg);
        }

        .custom-nav-item.is-open > .custom-nav-row > .custom-nav-flyout-toggle {
          transform: rotate(180deg);
        }
      }
    `;

    document.head.appendChild(this._styleTag);
  }
}