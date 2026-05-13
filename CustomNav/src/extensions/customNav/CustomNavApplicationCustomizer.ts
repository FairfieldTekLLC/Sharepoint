import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import SiteSearchService, { ISiteHit } from '../services/SiteSearchService';

export interface ICustomNavApplicationCustomizerProperties {}

interface ISiteNode {
  site: ISiteHit;
  children: ISiteNode[];
}

interface IExternalLink {
  title: string;
  url: string;
  target: '_self' | '_blank';
}

interface INavFilterProperties {
  hidePersonalSites?: boolean;
  hideTeamSites?: boolean;
  teamSitePathPrefixes?: string[];
}

interface ICachedMenuData {
  timestamp: number;
  sites: ISiteHit[];
}

interface ICustomNavWindow extends Window {
  __customNavInitialized?: boolean;
}

export default class CustomNavApplicationCustomizer
  extends BaseApplicationCustomizer<ICustomNavApplicationCustomizerProperties> {

  private static readonly MENU_CACHE_TTL_MS = 2 * 60 * 1000;

  private _placeholder: PlaceholderContent | undefined;
  private _fallbackHost: HTMLDivElement | undefined;
  private _navHost: HTMLDivElement | undefined;
  private _styleTag: HTMLStyleElement | undefined;
  private _menuInteractionsBound = false;

  private _onPlaceholdersChanged = (): void => {
    this._render().catch(() => undefined);
  };

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

  private _render = async (): Promise<void> => {
    let mountHost: HTMLElement;

    if (!this._placeholder) {
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
      if (this._fallbackHost) {
        this._fallbackHost.remove();
        this._fallbackHost = undefined;
      }
      mountHost = this._placeholder.domElement;
    } else {
      if (!this._fallbackHost) {
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
        await this.context.msGraphClientFactory.getClient('3'); // MSGraphClientV3 usage 【7-3cc616】
      this._renderNav(mountHost, graphClient);
    } catch (error) {
      // Keep fallback nav visible even if Graph cannot be initialized.
      console.warn('CustomNav: Graph client unavailable, keeping fallback nav.', error);
    }
  }

  private _renderNav(container: HTMLElement, graphClient?: MSGraphClientV3): void {
    if (!this._navHost) {
      this._navHost = document.createElement('div');
      this._navHost.id = 'custom-nav-root';
    }

    this._ensureStyles();

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
    externalLinks.forEach((link) => list.appendChild(this._renderExternalLinkItem(link)));

    this._navHost.appendChild(list);
    this._bindMenuInteractions();
  }

  private _getExternalLinks(): IExternalLink[] {
    const props = this.properties as { externalLinks?: Array<{ title?: string; url?: string; target?: string }> };
    const links = props?.externalLinks;

    if (!Array.isArray(links)) {
      return [];
    }

    return links
      .map((link) => {
        const target: '_self' | '_blank' = link.target === '_self' ? '_self' : '_blank';
        return {
          title: (link.title || '').trim(),
          url: (link.url || '').trim(),
          target
        };
      })
      .filter((link) => {
        if (!link.title || !link.url) {
          return false;
        }

        return /^https?:\/\//i.test(link.url);
      });
  }

  private _renderExternalLinkItem(link: IExternalLink): HTMLLIElement {
    const item = document.createElement('li');
    item.className = 'custom-nav-item custom-nav-item-external';

    const row = document.createElement('div');
    row.className = 'custom-nav-row';

    const anchor = document.createElement('a');
    anchor.className = 'custom-nav-link custom-nav-link-external';
    anchor.href = link.url;
    anchor.textContent = link.title;
    anchor.title = link.title;
    anchor.target = link.target;
    if (link.target === '_blank') {
      anchor.rel = 'noopener noreferrer';
    }

    row.appendChild(anchor);
    item.appendChild(row);
    return item;
  }

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

  private _getCacheKey(): string {
    const userKey = (this.context.pageContext.user.loginName || 'anonymous').toLowerCase();
    const siteKey = this.context.pageContext.site.absoluteUrl.toLowerCase();
    const filterKey = this._getFilterCacheKey();
    return `custom-nav-menu-cache::${siteKey}::${userKey}::${filterKey}`;
  }

  private _getFilterCacheKey(): string {
    const filterProps = this.properties as INavFilterProperties;
    const hidePersonal = filterProps.hidePersonalSites !== false;
    const hideTeam = filterProps.hideTeamSites !== false;
    const prefixes = this._getTeamPathPrefixes().join(',');
    return `${hidePersonal ? 'hp1' : 'hp0'}-${hideTeam ? 'ht1' : 'ht0'}-${prefixes}`;
  }

  private _shouldIncludeSite(url: string): boolean {
    const filterProps = this.properties as INavFilterProperties;
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

  private _getTeamPathPrefixes(): string[] {
    const filterProps = this.properties as INavFilterProperties;
    const configured = Array.isArray(filterProps.teamSitePathPrefixes)
      ? filterProps.teamSitePathPrefixes
      : ['/teams/'];

    return configured
      .map((prefix) => (prefix || '').trim().toLowerCase())
      .filter((prefix) => prefix.length > 0)
      .map((prefix) => (prefix.startsWith('/') ? prefix : `/${prefix}`));
  }

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

  private _siteLabel(site: ISiteHit): string {
    return site.displayName || site.name || site.webUrl || 'Site';
  }

  private _normalizeUrl(url: string): string {
    return url.replace(/\/$/, '').toLowerCase();
  }

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
      }

      .custom-nav-item {
        position: relative;
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
        padding: 6px 10px;
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
        padding: 6px 6px;
        line-height: 1;
        transition: transform 0.18s ease;
      }

      .custom-nav-flyout-toggle {
        font-size: 11px;
        padding-right: 10px;
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
        padding: 7px 12px;
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