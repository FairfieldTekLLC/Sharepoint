/**
 * CleanDocsWebPart.ts
 *
 * SPFx Web Part entry point for the CleanDocs web part.
 * Renders a sorted, clickable file list from a configured SharePoint
 * document library.  All files open in a new tab with navigation
 * interception disabled to avoid SharePoint client-side link hijacking.
 *
 * Property pane controls:
 *  - titleBarText  : Configurable heading shown above the document list.
 *  - libraryTitle  : Display name of the document library to query.
 *  - maxItems      : Upper limit on the number of files retrieved (1–500).
 */
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

/**
 * Strongly-typed property bag for the CleanDocs web part.
 * Values are persisted in the SharePoint page model and surfaced to the
 * property pane for editor configuration.
 */
export interface IYourWebPartProps {
  /** Heading text displayed in the CleanDocs title bar. Defaults to "Clean Documents". */
  titleBarText: string;

  /** Display name of the SharePoint document library to query. Defaults to "Documents". */
  libraryTitle: string;

  /** Maximum number of file items to retrieve from the library (property pane range 1–500). */
  maxItems: number;
}

/**
 * YourWebPart
 *
 * Extends BaseClientSideWebPart to mount the CleanDocs React component
 * and manage the property pane configuration.
 */

export default class YourWebPart extends BaseClientSideWebPart<IYourWebPartProps> {
  /**
   * render()
   *
   * Called by the SPFx framework whenever the web part needs to be (re-)rendered.
   * Resolves default values for unconfigured properties to guarantee a usable
   * initial state immediately after the web part is added to a page.
   * Mounts the CleanDocs React component into domElement.
   */
  public render(): void {
    const element: React.ReactElement<ICleanDocsProps> = React.createElement(CleanDocs, {
      // Absolute site URL provides the base for all REST API calls inside CleanDocs.
      siteUrl: this.context.pageContext.web.absoluteUrl,
      // Fall back to a sensible default if the property pane value is blank.
      titleBarText: this.properties.titleBarText || "Clean Documents",
      libraryTitle: this.properties.libraryTitle || "Documents",
      maxItems: this.properties.maxItems || 50,
      // Pass the SPHttpClient instance so CleanDocs can make authenticated REST calls.
      spHttpClient: this.context.spHttpClient
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /**
   * getPropertyPaneConfiguration()
   *
   * Defines the structure and controls rendered inside the property pane panel
   * shown to page editors.
   *
   * General group:
   *   titleBarText  — TextField: heading displayed in the component title bar.
   *   libraryTitle  — TextField: name of the document library to query.
   *   maxItems      — Slider:    maximum file count (range 1–500, step 1).
   */
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