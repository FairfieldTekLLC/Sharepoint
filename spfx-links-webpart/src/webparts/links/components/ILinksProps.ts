import { SPHttpClient } from '@microsoft/sp-http'; // SPHttpClient type 【5-d874ab】

export interface ILinksProps {
  title: string;
  webUrl: string;
  spHttpClient: SPHttpClient;

  listId: string;
  maxItems: number;
  showDescription: boolean;
  openInNewTab: boolean;
  category: string;

  description?: string;
  isDarkTheme?: boolean;
  environmentMessage?: string;
  hasTeamsContext?: boolean;
  userDisplayName?: string;
}