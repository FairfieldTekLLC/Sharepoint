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
} from '@microsoft/sp-property-pane'; // property pane fields 【3-3d17bb】【9-e1c3b3】

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http'; // SPHttpClient for REST 【5-d874ab】【6-7acd82】
import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane'; // dropdown option shape 【4-0b5d9e】

import Links from './components/Links';
import { ILinksProps } from './components/ILinksProps';

export interface ILinksWebPartProps {
  title: string;
  description?: string;
  listId: string;
  maxItems: number;
  showDescription: boolean;
  openInNewTab: boolean;
  category: string;
}

interface IODataList {
  Id: string;
  Title: string;
  Hidden: boolean;
}

export default class LinksWebPart extends BaseClientSideWebPart<ILinksWebPartProps> {
  private _listOptions: IPropertyPaneDropdownOption[] = [];
  private _listsLoaded: boolean = false;

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

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  private async _loadLists(): Promise<void> {
    if (this._listsLoaded) return;

    // SharePoint REST: get non-hidden lists 【7-1f2a7b】
    const endpoint =
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists` +
      `?$select=Id,Title,Hidden&$filter=Hidden eq false&$orderby=Title`;

    // SPHttpClient is used for SharePoint REST 【5-d874ab】【6-7acd82】
    const res: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          // Microsoft shows accept header usage, including odata.metadata=none 【6-7acd82】
          'accept': 'application/json;odata.metadata=none'
        }
      }
    );

    const json = await res.json();
    const lists: IODataList[] = json.value ?? [];

    // Dropdown options are key/text pairs 【4-0b5d9e】
    this._listOptions = lists.map(l => ({ key: l.Id, text: l.Title }));
    this._listsLoaded = true;

    // If dropdown shows empty initially, refresh property pane UI (common pattern) 【10-ad1554】【11-7fcf35】
    this.context.propertyPane.refresh();
  }

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