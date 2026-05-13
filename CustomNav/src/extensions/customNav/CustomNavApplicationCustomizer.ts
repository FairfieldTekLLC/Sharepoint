import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import SiteSearchService, { ISiteHit } from '../services/SiteSearchService';

export interface ICustomNavApplicationCustomizerProperties {}

export default class CustomNavApplicationCustomizer
  extends BaseApplicationCustomizer<ICustomNavApplicationCustomizerProperties> {

  private _placeholder: PlaceholderContent | undefined;
  private _fallbackHost: HTMLDivElement | undefined;
  private _navHost: HTMLDivElement | undefined;

  private _onPlaceholdersChanged = (): void => {
    this._render().catch(() => undefined);
  };

  public onInit(): Promise<void> {
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

    this._navHost.innerHTML = '';
    this._navHost.style.display = 'flex';
    this._navHost.style.alignItems = 'center';
    this._navHost.style.gap = '12px';
    this._navHost.style.padding = '8px 12px';
    this._navHost.style.background = '#0b1220';
    this._navHost.style.color = '#fff';
    this._navHost.style.fontFamily = 'Segoe UI, Arial, sans-serif';

    const title = document.createElement('strong');
    title.textContent = 'CustomNav';
    this._navHost.appendChild(title);

    if (!graphClient) {
      const unavailable = document.createElement('div');
      unavailable.textContent = 'Search is temporarily unavailable. Basic navigation is still active.';
      unavailable.style.marginLeft = 'auto';
      unavailable.style.fontSize = '13px';
      unavailable.style.opacity = '0.9';
      this._navHost.appendChild(unavailable);
      container.replaceChildren(this._navHost);
      return;
    }

    const service = new SiteSearchService(graphClient);
    const input = document.createElement('input');
    input.placeholder = 'Search sites you can access';
    input.style.width = '320px';
    input.style.padding = '6px 8px';
    input.style.borderRadius = '8px';
    input.style.border = '1px solid #cbd5e1';

    const button = document.createElement('button');
    button.textContent = 'Search';
    button.style.padding = '6px 10px';
    button.style.borderRadius = '8px';
    button.style.border = '1px solid #334155';
    button.style.cursor = 'pointer';

    const results = document.createElement('div');
    results.style.marginLeft = 'auto';
    results.style.display = 'flex';
    results.style.gap = '8px';
    results.style.flexWrap = 'wrap';
    results.style.maxWidth = '700px';

    button.addEventListener('click', () => {
      const query = input.value || '';
      service.searchSites(query, 0, 8)
        .then((r: { items: ISiteHit[] }) => {
          results.innerHTML = '';
          r.items.slice(0, 8).forEach((site) => {
            const link = document.createElement('a');
            link.href = site.webUrl || '#';
            link.textContent = site.displayName || site.name || site.webUrl || 'Open site';
            link.style.color = '#93c5fd';
            link.style.textDecoration = 'none';
            results.appendChild(link);
          });
        })
        .catch(() => {
          results.innerHTML = '';
          const error = document.createElement('span');
          error.textContent = 'Search unavailable';
          error.style.opacity = '0.9';
          results.appendChild(error);
        });
    });

    this._navHost.appendChild(input);
    this._navHost.appendChild(button);
    this._navHost.appendChild(results);
    container.replaceChildren(this._navHost);
  }
}