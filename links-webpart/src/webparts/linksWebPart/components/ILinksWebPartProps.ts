import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ILinksWebPartProps {
  context: WebPartContext;
  listTitle: string;
  maxItems: number;
  openInNewTab: boolean;
  showTopText: boolean;
  showDescription: boolean;
  backgroundColor: string;
  titleBarColor: string;
  linkTextColor: string;
}