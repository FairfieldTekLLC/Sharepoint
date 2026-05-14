/**
 * Type declaration for the generated localization string bundle used by the
 * ChatBot web part. These keys map to entries in the locale-specific resource
 * files emitted by the SPFx build pipeline.
 */
declare interface IChatBotWebPartWebPartStrings {
  /** Property pane description shown at the top of the edit panel. */
  PropertyPaneDescription: string;
  /** Label for the default property pane group. */
  BasicGroupName: string;
  /** Legacy description field label. */
  DescriptionFieldLabel: string;
  /** Environment label used when running locally in SharePoint. */
  AppLocalEnvironmentSharePoint: string;
  /** Environment label used when running locally in Teams. */
  AppLocalEnvironmentTeams: string;
  /** Environment label used when running locally in Office. */
  AppLocalEnvironmentOffice: string;
  /** Environment label used when running locally in Outlook. */
  AppLocalEnvironmentOutlook: string;
  /** Environment label used when running in hosted SharePoint. */
  AppSharePointEnvironment: string;
  /** Environment label used when running in a Teams tab. */
  AppTeamsTabEnvironment: string;
  /** Environment label used when running in hosted Office. */
  AppOfficeEnvironment: string;
  /** Environment label used when running in hosted Outlook. */
  AppOutlookEnvironment: string;
  /** Fallback environment label for unknown hosts. */
  UnknownEnvironment: string;
}

/** Module declaration that lets TypeScript import the generated string bundle. */
declare module 'ChatBotWebPartWebPartStrings' {
  const strings: IChatBotWebPartWebPartStrings;
  export = strings;
}
