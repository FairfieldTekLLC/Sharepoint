/**
 * ICleanDocsProps.ts
 *
 * Defines the props contract for the CleanDocs React component.
 * All values are resolved by the SPFx web part class (YourWebPart) from the
 * property pane configuration and the active page context.
 */
import { SPHttpClient } from "@microsoft/sp-http";

/**
 * Props accepted by the CleanDocs functional component.
 *
 * CleanDocs renders a sorted, clickable list of files from a specified
 * SharePoint document library, opening each document in a new tab.
 */
export interface ICleanDocsProps {
  /**
   * Absolute URL of the SharePoint site that contains the document library.
   * Typically sourced from `this.context.pageContext.web.absoluteUrl` in the
   * web part class. Used as the base for all REST API calls.
   */
  siteUrl: string;

  /**
   * Text displayed in the component's title bar.
   * Configurable from the property pane; defaults to "Clean Documents" if blank.
   */
  titleBarText: string;

  /**
   * Display name of the SharePoint document library to query.
   * Must match the library's Title exactly (case-insensitive on most tenants).
   * Configurable from the property pane; defaults to "Documents" if blank.
   */
  libraryTitle: string;

  /**
   * The SPFx SPHttpClient instance from the current context.
   * Used to make authenticated REST calls to the SharePoint REST API.
   * Passed through from the web part so the component does not need its own context.
   */
  spHttpClient: SPHttpClient;

  /**
   * Maximum number of file items to retrieve from the library.
   * Configurable from the property pane (range 1–500); defaults to 50.
   * Increasing this value raises the REST payload size.
   */
  maxItems: number;
}
