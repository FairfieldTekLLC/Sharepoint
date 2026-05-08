import * as React from "react";
import * as ReactDom from "react-dom";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from "@microsoft/sp-webpart-base";
import { CleanDocs } from "./components/CleanDocs";
import { ICleanDocsProps } from "./components/ICleanDocsProps";

export interface IYourWebPartProps {
  titleBarText: string;
  libraryTitle: string;
  maxItems: number;
}

export default class YourWebPart extends BaseClientSideWebPart<IYourWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ICleanDocsProps> = React.createElement(CleanDocs, {
      siteUrl: this.context.pageContext.web.absoluteUrl,
      titleBarText: this.properties.titleBarText || "Clean Documents",
      libraryTitle: this.properties.libraryTitle || "Documents",
      maxItems: this.properties.maxItems || 50,
      spHttpClient: this.context.spHttpClient
    });

    ReactDom.render(element, this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Web Part Settings" },
          groups: [
            {
              groupName: "General",
              groupFields: [
                PropertyPaneTextField("titleBarText", {
                  label: "Title Bar Text"
                }),
                PropertyPaneTextField("libraryTitle", {
                  label: "Document Library Title"
                }),
                PropertyPaneSlider("maxItems", {
                  label: "Max Items",
                  min: 1,
                  max: 500,
                  step: 1
                })
              ]
            }
          ]
        }
      ]
    };
  }
}