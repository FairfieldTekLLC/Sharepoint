import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IChatBotWebPartProps {
  context: WebPartContext;
  agentUrl?: string;
  tokenEndpointUrl: string;
  title: string;
  height: number;
  width: number;
}

