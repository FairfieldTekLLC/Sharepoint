import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'CustomNavOnPremiseWebPartStrings';
import CustomNavOnPremise from './components/CustomNavOnPremise';
import { ICustomNavOnPremiseProps } from './components/ICustomNavOnPremiseProps';

export interface ICustomNavOnPremiseWebPartProps {
  description: string;
  customNavigationJson: string;
}

export default class CustomNavOnPremiseWebPart extends BaseClientSideWebPart<ICustomNavOnPremiseWebPartProps> {

  private _hideClassicTitleChrome(): void {
    const webPartChrome: HTMLElement | null = this.domElement.closest('.ms-webpart-chrome') as HTMLElement;

    if (!webPartChrome) {
      return;
    }

    const titleSelectors: string[] = [
      '.ms-webpart-chrome-title',
      '.ms-webpart-titleText',
      '.ms-webpart-titleText > span',
      '.js-webpart-titleCell'
    ];

    titleSelectors.forEach((selector: string) => {
      const titleElement: HTMLElement | null = webPartChrome.querySelector(selector) as HTMLElement;
      if (titleElement) {
        titleElement.style.display = 'none';
      }
    });
  }

  public render(): void {
    const pageContext: any = this.context.pageContext as any;
    const webUrl: string = pageContext && pageContext.web && pageContext.web.absoluteUrl
      ? pageContext.web.absoluteUrl
      : window.location.origin;
    const siteUrl: string = pageContext && pageContext.site && pageContext.site.absoluteUrl
      ? pageContext.site.absoluteUrl
      : webUrl;

    const element: React.ReactElement<ICustomNavOnPremiseProps > = React.createElement(
      CustomNavOnPremise,
      {
        description: this.properties.description,
        customNavigationJson: this.properties.customNavigationJson,
        siteUrl: siteUrl,
        currentWebUrl: webUrl
      }
    );

    ReactDom.render(element, this.domElement);
    this._hideClassicTitleChrome();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('customNavigationJson', {
                  label: strings.CustomNavigationJsonFieldLabel,
                  multiline: true,
                  rows: 10
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
