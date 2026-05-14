/**
 * LinksWebPartWebPart.ts  (links-webpart)
 *
 * SPFx Web Part entry point for the LinksWebPart web part.
 * Renders a categorised, collapsible list of SharePoint list hyperlinks with
 * fully configurable colours and display options.
 *
 * Property pane groups:
 *  Data    — listTitle, maxItems
 *  Display — backgroundColor, titleBarColor, linkTextColor,
 *             openInNewTab, showTopText, showDescription
 */
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



/**
 * Strongly-typed property bag for the LinksWebPart web part.
 * Values are persisted in the SharePoint page model and surfaced via the
 * property pane for editor configuration.
 */
export interface ILinksWebPartWebPartProps {
  /** Display name of the SharePoint list that contains link items. */
  listTitle: string;

  /** Maximum number of link items to retrieve (range 1–200). */
  maxItems: number;

  /** When true, links open in a new tab; false opens in the current tab. */
  openInNewTab: boolean;

  /** When true, shows the static IT Hotline callout box above the link list. */
  showTopText: boolean;

  /** When true, renders each link item's Description field beneath the anchor. */
  showDescription: boolean;

  /** CSS color value for the card background. Applied as --links-background. */
  backgroundColor: string;

  /** CSS color value for category header bars. Applied as --links-titlebar-bg. */
  titleBarColor: string;

  /** CSS color value for anchor text. Applied as --links-text-color. */
  linkTextColor: string;
}

export default class LinksWebPartWebPart extends BaseClientSideWebPart<ILinksWebPartWebPartProps> {

  /** Tracks whether the current SharePoint theme uses an inverted (dark) colour palette. */
  private _isDarkTheme: boolean = false;

  /** Human-readable environment label resolved from the Teams/Office host context. */
  private _environmentMessage: string = '';

  /**
   * render()
   *
   * Called by the SPFx framework whenever the web part needs to be (re-)rendered.
   * Applies default values for any unconfigured properties before mounting the
   * LinksWebPart React component, ensuring a usable out-of-box experience.
   */
  public render(): void {
  // Apply safe defaults if properties have not been configured by the editor.
  if (!this.properties.listTitle) this.properties.listTitle = "Links";
  if (!this.properties.maxItems) this.properties.maxItems = 12;
  if (this.properties.openInNewTab === undefined) this.properties.openInNewTab = true;
  if (this.properties.showTopText === undefined) this.properties.showTopText = true;
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
      showTopText: this.properties.showTopText,
      showDescription: this.properties.showDescription,
      backgroundColor: this.properties.backgroundColor,
      titleBarColor: this.properties.titleBarColor,
      linkTextColor: this.properties.linkTextColor
    }
  );

  ReactDom.render(element, this.domElement);
}
 

  /**
   * onInit()
   *
   * SPFx lifecycle hook called once before the first render.
   * Resolves the environment message string (used for Teams / Office host detection)
   * and stores it for potential display or diagnostic use.
   */
  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  /**
   * _getEnvironmentMessage()
   *
   * Determines the human-readable environment context label.
   * When running inside Teams, Office, or Outlook, the label reflects both
   * the host app and whether the bundle is served from localhost (dev mode)
   * or production.
   *
   * @returns A promise resolving to the environment description string.
   */
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

  /**
   * onThemeChanged()
   *
   * SPFx lifecycle hook called when the page theme changes (e.g., user switches
   * to a dark variant).  Updates the local dark-theme flag and injects semantic
   * colour tokens as CSS custom properties on domElement so that SCSS rules can
   * reference them with `var(--bodyText)`, `var(--link)`, etc.
   *
   * @param currentTheme - The new theme object, or undefined if the theme was reset.
   */
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

  /**
   * onDispose()
   *
   * SPFx lifecycle hook called when the web part is removed from the page.
   * Unmounts the React component tree to release all event handlers and
   * prevent memory leaks.
   */
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /**
   * dataVersion
   *
   * Semantic version of the serialised property bag.
   * Increment when making breaking changes to ILinksWebPartWebPartProps so the
   * SPFx framework can detect and migrate stale page data.
   */
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * getPropertyPaneConfiguration()
   *
   * Defines the structure and controls rendered in the property pane panel.
   *
   * Data group:
   *   listTitle        — TextField: SharePoint list name.
   *   maxItems         — Slider:    max links to show (1–200).
   * Display group:
   *   backgroundColor  — ColorPicker: card background.
   *   titleBarColor    — ColorPicker: category header bar colour.
   *   linkTextColor    — ColorPicker: anchor text colour.
   *   openInNewTab     — Toggle:      open links in new tab.
   *   showTopText      — Toggle:      show IT Hotline callout.
   *   showDescription  — Toggle:      show item descriptions.
   */
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
              PropertyPaneToggle("showTopText", {
                label: "Show top text"
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
