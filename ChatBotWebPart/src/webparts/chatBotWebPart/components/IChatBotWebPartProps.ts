/**
 * IChatBotWebPartProps.ts
 *
 * Defines the props contract for the ChatBotWebPart (ChatWidget) React component.
 * All values are supplied by the SPFx web part class (ChatBotWebPartWebPart) from the
 * property pane configuration and the active WebPartContext.
 */
import { WebPartContext } from "@microsoft/sp-webpart-base";

/**
 * Props accepted by the ChatBotWebPart functional component.
 *
 * The component supports two mutually exclusive rendering modes:
 *  1. iframe mode   — set `agentUrl` to a published Copilot Studio or Azure Bot
 *                     Service URL; the widget embeds it in a sandboxed <iframe>.
 *  2. WebChat mode  — leave `agentUrl` blank and provide `tokenEndpointUrl`;
 *                     the component bootstraps the BotFramework WebChat SDK and
 *                     performs an AAD token exchange for SSO.
 */
export interface IChatBotWebPartProps {
  /**
   * The SPFx WebPartContext for the current page.
   * Provides access to `aadTokenProviderFactory` (used for acquiring user tokens
   * during the Direct Line OAuth token-exchange flow) and user display name.
   */
  context: WebPartContext;

  /**
   * Published Copilot Studio or Azure Bot Service iframe URL.
   * When non-empty the widget renders this URL in an <iframe>, bypassing the
   * BotFramework WebChat SDK entirely.
   * Optional — omit or leave blank to activate the WebChat / Direct Line path.
   */
  agentUrl?: string;

  /**
   * Direct Line token endpoint URL (e.g. a Copilot Studio "Generate token" endpoint).
   * Required when `agentUrl` is not set.
   * The component issues a GET request here to obtain a short-lived conversation
   * token before initialising the BotFramework WebChat renderer.
   */
  tokenEndpointUrl: string;

  /** Display label shown in the chat widget header bar. */
  title: string;

  /** Height of the rendered chat widget in pixels. */
  height: number;

  /** Width of the rendered chat widget in pixels. */
  width: number;
}

