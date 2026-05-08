import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import {
  PropertyFieldColorPicker,
  PropertyFieldColorPickerStyle
} from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'LinksWebPartWebPartStrings';
import LinksWebPart from './components/LinksWebPart';



export interface ILinksWebPartWebPartProps {
  listTitle: string;
  maxItems: number;
  openInNewTab: boolean;
  showDescription: boolean;
  backgroundColor: string;
  titleBarColor: string;
  linkTextColor: string;
}

export default class LinksWebPartWebPart extends BaseClientSideWebPart<ILinksWebPartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
  // Defaults (safe if user hasn't configured yet)
  if (!this.properties.listTitle) this.properties.listTitle = "Links";
  if (!this.properties.maxItems) this.properties.maxItems = 12;
  if (this.properties.openInNewTab === undefined) this.properties.openInNewTab = true;
  if (this.properties.showDescription === undefined) this.properties.showDescription = true;
  if (!this.properties.backgroundColor) this.properties.backgroundColor = '#f7f7f7';
  if (!this.properties.titleBarColor) this.properties.titleBarColor = '#d1d3aa';
  if (!this.properties.linkTextColor) this.properties.linkTextColor = '#0e0e0e';

  const element: React.ReactElement = React.createElement(
    LinksWebPart, // your component import
    {
      context: this.context,
      listTitle: this.properties.listTitle,
      maxItems: this.properties.maxItems,
      openInNewTab: this.properties.openInNewTab,
      showDescription: this.properties.showDescription,
      backgroundColor: this.properties.backgroundColor,
      titleBarColor: this.properties.titleBarColor,
      linkTextColor: this.properties.linkTextColor
    }
  );

  ReactDom.render(element, this.domElement);
}
 

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  return {
    pages: [
      {
        header: { description: "Links Web Part Settings" },
        groups: [
          {
            groupName: "Data",
            groupFields: [
              PropertyPaneTextField("listTitle", {
                label: "List title",
                description: "Name of the SharePoint list that contains your links"
              }),
              PropertyPaneSlider("maxItems", {
                label: "Max links to show",
                min: 1,
                max: 200,
                value: this.properties.maxItems || 12
              })
            ]
          },
          {
            groupName: "Display",
            groupFields: [
              PropertyFieldColorPicker('backgroundColor', {
                label: 'Background color',
                selectedColor: this.properties.backgroundColor,
                onPropertyChange: this.onPropertyPaneFieldChanged,
                properties: this.properties,
                disabled: false,
                debounce: 300,
                style: PropertyFieldColorPickerStyle.Inline,
                key: 'backgroundColorField'
              }),
              PropertyFieldColorPicker('titleBarColor', {
                label: 'Title bar color',
                selectedColor: this.properties.titleBarColor,
                onPropertyChange: this.onPropertyPaneFieldChanged,
                properties: this.properties,
                disabled: false,
                debounce: 300,
                style: PropertyFieldColorPickerStyle.Inline,
                key: 'titleBarColorField'
              }),
              PropertyFieldColorPicker('linkTextColor', {
                label: 'Link text color',
                selectedColor: this.properties.linkTextColor,
                onPropertyChange: this.onPropertyPaneFieldChanged,
                properties: this.properties,
                disabled: false,
                debounce: 300,
                style: PropertyFieldColorPickerStyle.Inline,
                key: 'linkTextColorField'
              }),
              PropertyPaneToggle("openInNewTab", {
                label: "Open links in a new tab"
              }),
              PropertyPaneToggle("showDescription", {
                label: "Show description"
              })
            ]
          }
        ]
      }
    ]
  };
}

}
