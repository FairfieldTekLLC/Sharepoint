import { SPHttpClient } from "@microsoft/sp-http";

export interface ICleanDocsProps {
  siteUrl: string;
  titleBarText: string;
  libraryTitle: string;
  spHttpClient: SPHttpClient;
  maxItems: number;
}
