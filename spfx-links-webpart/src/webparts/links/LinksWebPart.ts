/**
 * LinksWebPart.ts  (spfx-links-webpart)
 *
 * SPFx Web Part entry point for the spfx-links-webpart.
 * Renders a flat, ordered list of time-gated hyperlinks sourced from a
 * SharePoint list selected by the page editor via a dropdown populated
 * dynamically at property pane open time.
 *
 * Property pane groups:
 *  Data    — listId (dropdown), category (optional filter), maxItems
 *  Display — title, showDescription, openInNewTab
 */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';

import Links from './components/Links';
import { ILinksProps } from './components/ILinksProps';

/**
 * Strongly-typed property bag for the spfx-links-webpart web part.
 * Values are persisted in the SharePoint page model and surfaced via the
 * property pane for editor configuration.
 */
export interface ILinksWebPartProps {
  /** Display title rendered as a heading above the link list. */
  title: string;

  /** Optional secondary description text (currently unused by the Links component). */
  description?: string;

  /**
   * GUID of the SharePoint list that contains the link items.
   * Using the GUID (rather than Title) avoids breakage on list renames.
   * Populated from the property pane dropdown that enumerates visible lists.
   */
  listId: string;

  /** Maximum number of items to retrieve (range 1\u201350). */
  maxItems: number;

  /** When true, renders item descriptions beneath each anchor. */
  showDescription: boolean;

  /** When true, each link opens in a new browser tab. */
  openInNewTab: boolean;

  /**
   * Optional category filter string.
   * When non-empty, only items in the specified category are fetched.
   */
  category: string;
}

/**
 * Raw shape of a SharePoint list object returned by the REST Lists endpoint.
 * Only the fields selected by `_loadLists` are present.
 */
interface IODataList {
  /** GUID identifier of the list. */
  Id: string;

  /** Display title of the list. */
  Title: string;

  /**
   * Whether the list is hidden from users.
   * Hidden lists are excluded from the dropdown to avoid confusing editors.
   */
  Hidden: boolean;
}

/**
 * LinksWebPart
 *
 * Extends BaseClientSideWebPart to mount the Links React component and
 * dynamically populate the list-selection dropdown in the property pane.
 */
export default class LinksWebPart extends BaseClientSideWebPart<ILinksWebPartProps> {
  /**
   * Dropdown options for the list-selection control.
   * Populated by `_loadLists` during `onInit`; refreshed into the property pane
   * via `this.context.propertyPane.refresh()` once loading completes.
   */
  private _listOptions: IPropertyPaneDropdownOption[] = [];

  /**
   * Flag that prevents `_loadLists` from re-fetching once it has completed.
   * Also used to control the `disabled` state of the list dropdown while loading.
   */
  private _listsLoaded: boolean = false;

  /**
   * onInit()
   *
   * SPFx lifecycle hook called once before the first render.
   * Pre-loads the available SharePoint lists so the property pane dropdown is
   * ready when the editor opens the panel.  Failures are swallowed to ensure
   * the web part remains usable (the dropdown will be empty, not broken).
   */
  protected async onInit(): Promise<void> {
    await super.onInit();

    try {
      await this._loadLists();
    } catch {
      // Keep the web part usable even when list discovery fails.
      this._listsLoaded = true;
      this._listOptions = [];
    }
  }

  /**
   * render()
   *
   * Called by the SPFx framework whenever the web part needs to be (re-)rendered.
   * Forwards all configured properties and context values to the Links React
   * component, applying null-coalescing defaults for unconfigured properties.
   */
  public render(): void {
    const element: React.ReactElement<ILinksProps> = React.createElement(Links, {
      title: this.properties.title ?? '',
      webUrl: this.context.pageContext.web.absoluteUrl,
      spHttpClient: this.context.spHttpClient,
      listId: this.properties.listId ?? '',
      maxItems: this.properties.maxItems ?? 10,
      showDescription: this.properties.showDescription ?? true,
      openInNewTab: this.properties.openInNewTab ?? true,
      category: this.properties.category ?? '',
      description: this.properties.description,
      isDarkTheme: false,
      environmentMessage: this.context.isServedFromLocalhost ? 'Local' : 'Remote',
      hasTeamsContext: !!this.context.sdks?.microsoftTeams,
      userDisplayName: this.context.pageContext.user?.displayName ?? ''
    });

    ReactDom.render(element, this.domElement);
  }

  /**
   * onDispose()
   *
   * SPFx lifecycle hook called when the web part is removed from the page.
   * Unmounts the React component tree to release event listeners and
   * prevent memory leaks.
   */
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /**
   * _loadLists()
   *
   * Fetches all non-hidden SharePoint lists from the current site via the REST
   * API and populates `_listOptions` for the property pane list dropdown.
   *
   * Guard: exits immediately if lists have already been loaded to prevent
   * redundant network calls (e.g., when the property pane is reopened).
   *
   * After loading, calls `this.context.propertyPane.refresh()` so the now-populated
   * dropdown becomes visible without requiring the editor to reopen the panel.
   */
  private async _loadLists(): Promise<void> {
    // Return early if we have already fetched lists for this web part instance.
    if (this._listsLoaded) return;

    // Build the REST endpoint; $filter=Hidden eq false excludes system/internal lists.
    const endpoint =
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists` +
      `?$select=Id,Title,Hidden&$filter=Hidden eq false&$orderby=Title`;

    // SPHttpClient handles authentication transparently via the SPFx runtime.
    // OData metadata=none reduces the response payload size.
    const res: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'accept': 'application/json;odata.metadata=none'
        }
      }
    );

    const json = await res.json();
    const lists: IODataList[] = json.value ?? [];

    // Map each list to a { key, text } pair expected by PropertyPaneDropdown.
    this._listOptions = lists.map(l => ({ key: l.Id, text: l.Title }));
    this._listsLoaded = true;

    // Trigger a property pane re-render so the now-populated dropdown is visible.
    this.context.propertyPane.refresh();
  }

  /**
   * getPropertyPaneConfiguration()
   *
   * Defines the structure and controls rendered in the property pane panel.
   *
   * Data group:
   *   listId    — Dropdown: select from discovered non-hidden lists (disabled while loading).
   *   category  — TextField: optional category filter (empty = all categories).
   *   maxItems  — Slider:    maximum result count (1–50).
   * Display group:
   *   title           — TextField: heading shown above the link list.
   *   showDescription — Toggle:    show item description beneath each link.
   *   openInNewTab    — Toggle:    open links in new tab.
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Links web part settings' },
          groups: [
            {
              groupName: 'Data',
              groupFields: [
                PropertyPaneDropdown('listId', {
                  label: 'Select a links list',
                  options: this._listOptions,
                  disabled: !this._listsLoaded
                }),
                PropertyPaneTextField('category', {
                  label: 'Category filter (optional)'
                }),
                PropertyPaneSlider('maxItems', {
                  label: 'Max items',
                  min: 1,
                  max: 50,
                  value: this.properties.maxItems ?? 10
                })
              ]
            },
            {
              groupName: 'Display',
              groupFields: [
                PropertyPaneTextField('title', { label: 'Title' }),
                PropertyPaneToggle('showDescription', { label: 'Show description' }),
                PropertyPaneToggle('openInNewTab', { label: 'Open in new tab' })
              ]
            }
          ]
        }
      ]
    };
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}