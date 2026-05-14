/**
 * ChatBotWebPart.tsx  (ChatWidget)
 *
 * React functional component that renders an embedded AI chat widget inside
 * a SharePoint page.  Supports two rendering modes:
 *
 *  1. iframe mode  — When `agentUrl` is non-empty the agent URL is loaded in a
 *                    sandboxed <iframe>.  No BotFramework SDK required.
 *  2. WebChat mode — When `agentUrl` is absent, the BotFramework WebChat SDK
 *                    is loaded from CDN and rendered against a Direct Line token.
 *                    An AAD token exchange middleware intercepts OAuth card
 *                    activities to enable silent SSO for the signed-in user.
 *
 * The widget can be minimised/expanded via a toggle button in the header bar.
 */
import * as React from "react";
import { AadTokenProvider } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

/**
 * Content type identifier for Bot Framework OAuth card attachments.
 * Used by the WebChat store middleware to detect and intercept sign-in prompts.
 */
const OAuthCardContentType = "application/vnd.microsoft.card.oauth";

/**
 * Props for the ChatWidget component.
 * Extends the base IChatBotWebPartProps with a required WebPartContext.
 */
export interface IChatWidgetProps {
  /** The SPFx WebPartContext; provides AAD token provider and user information. */
  context: WebPartContext;

  /**
   * Published Copilot Studio or Azure Bot Service iframe URL.
   * When non-empty the component skips WebChat SDK initialization and renders
   * the URL directly inside an <iframe>.
   */
  agentUrl?: string;

  /**
   * Direct Line token endpoint URL.
   * Used when `agentUrl` is not provided; the component GETs a conversation
   * token from this URL before rendering the BotFramework WebChat SDK.
   */
  tokenEndpointUrl: string;

  /** Text label shown in the widget header bar. */
  title: string;

  /** Height of the widget in pixels. */
  height: number;

  /** Width of the widget in pixels. */
  width: number;
}

/**
 * Augment the global Window type so TypeScript recognises `window.WebChat`
 * after the BotFramework WebChat bundle is dynamically loaded from CDN.
 */
declare global {
  interface Window {
    WebChat: any;
  }
}

/**
 * ChatWidget
 *
 * Main exported component rendered by ChatBotWebPartWebPart.
 * Manages its own open/minimised state and the lazy initialisation of the
 * BotFramework WebChat SDK.
 *
 * @param props - Configuration props forwarded from the SPFx web part class.
 */
export default function ChatWidget(props: IChatWidgetProps) {
  /** Controls whether the widget body is visible (true) or minimised (false). */
  const [open, setOpen] = React.useState(true);

  /**
   * Tracks whether the BotFramework WebChat SDK has been rendered.
   * Prevents re-initialisation on subsequent re-renders.
   */
  const [initialized, setInitialized] = React.useState(false);

  /** Ref to the <div> that WebChat.renderWebChat() mounts the chat UI into. */
  const hostRef = React.useRef<HTMLDivElement>(null);

  /** Trim whitespace so empty strings are treated correctly throughout the component. */
  const trimmedAgentUrl = (props.agentUrl || "").trim();

  /** True when the user has configured an agent URL; activates iframe rendering mode. */
  const useIframeAgent = trimmedAgentUrl.length > 0;

  /**
   * initChat
   *
   * Async function (memoised with useCallback) that bootstraps the BotFramework
   * WebChat SDK and renders it into `hostRef.current`.
   *
   * Execution steps:
   *  1. Guard: skip if using iframe mode, already initialized, or DOM ref is not ready.
   *  2. Guard: validate that a token endpoint URL has been configured.
   *  3. Dynamically load the BotFramework WebChat bundle from the public CDN if it is
   *     not already present on the page (avoids duplicate script loading).
   *  4. Obtain a Direct Line conversation token from the configured token endpoint.
   *  5. Acquire an AAD token provider for the current SharePoint user.
   *  6. Create a WebChat store with middleware that intercepts OAuth card activities
   *     and performs a silent token exchange using the user's AAD token, enabling SSO.
   *  7. Render the WebChat component into the host <div>.
   */
  const initChat = React.useCallback(async () => {
    // Skip initialisation when iframe mode is active (no SDK needed).
    if (useIframeAgent) return;
    // Prevent double-initialization on re-renders.
    if (initialized || !hostRef.current) return;

    if (!props.tokenEndpointUrl || !props.tokenEndpointUrl.trim()) {
      throw new Error("Set Agent URL or Token Endpoint URL in web part properties.");
    }

    // 1) Load BotFramework WebChat script (or bundle import instead)
    if (!window.WebChat) {
      await new Promise<void>((resolve, reject) => {
        const s = document.createElement("script");
        s.src = "https://cdn.botframework.com/botframework-webchat/latest/webchat.js";
        s.onload = () => resolve();
        s.onerror = () => reject(new Error("Failed to load WebChat"));
        document.head.appendChild(s);
      });
    }

    // 2) Get Direct Line token from Copilot Studio token endpoint (GET returns token)
    const dlRes = await fetch(props.tokenEndpointUrl, { method: "GET" });
    if (!dlRes.ok) throw new Error(`Token endpoint failed: ${dlRes.status}`);
    const { token } = await dlRes.json(); // token, expires_in, conversationId 【4-95e277】

    // 3) Prepare AadTokenProvider (signed-in SharePoint user token acquisition)
    const aadTokenProvider: AadTokenProvider = await props.context.aadTokenProviderFactory.getTokenProvider();

    async function exchangeTokenAsync(resourceUri: string): Promise<string> {
      // Acquire a user token for the resource URI requested by the OAuth card
      return aadTokenProvider.getToken(resourceUri);
    }

    // 4) WebChat store middleware: intercept OAuth card and do signin/tokenExchange
    const store = window.WebChat.createStore(
      {},
      ({ dispatch }: any) => (next: any) => async (action: any) => {
        if (action.type === "DIRECT_LINE/INCOMING_ACTIVITY") {
          const activity = action.payload?.activity;

          const attachment = activity?.attachments?.[0];
          if (attachment?.contentType === OAuthCardContentType) {
            try {
              const content = attachment.content;

              // OAuth card contains tokenExchangeResource.uri, tokenExchangeResource.id, and connectionName 【2-d4d5fa】
              const resourceUri = content?.tokenExchangeResource?.uri;
              const exchangeId = content?.tokenExchangeResource?.id;
              const connectionName = content?.connectionName;

              if (resourceUri && exchangeId && connectionName) {
                const userAccessToken = await exchangeTokenAsync(resourceUri);

                // Post signin/tokenExchange invoke activity to Direct Line 【2-d4d5fa】
                await dispatch({
                  type: "DIRECT_LINE/POST_ACTIVITY",
                  payload: {
                    activity: {
                      type: "invoke",
                      name: "signin/tokenExchange",
                      from: { id: "spfx-user", name: props.context.pageContext.user.displayName },
                      value: {
                        id: exchangeId,
                        connectionName,
                        token: userAccessToken
                      }
                    }
                  },
                  meta: { method: "keyboard" }
                });

                // Swallow OAuth card so user never sees it (SSO experience) 【2-d4d5fa】
                return;
              }
            } catch (e) {
              // If token exchange fails, fall through so the OAuth card shows.
              console.error("Token exchange failed; showing OAuth card.", e);
            }
          }
        }

        return next(action);
      }
    );

    // 5) Render WebChat
    const directLine = window.WebChat.createDirectLine({ token });
    window.WebChat.renderWebChat(
      {
        directLine,
        store,
        styleOptions: {
          hideUploadButton: true
        }
      },
      hostRef.current
    );

    setInitialized(true);
  }, [initialized, props, useIframeAgent]);

  // Initialize only when user opens widget (user gesture helps popup/token flows) 【9-9f860e】
  React.useEffect(() => {
    if (open) {
      initChat().catch(err => console.error(err));
    }
  }, [open, initChat]);

  return (
    <div
      style={{
        width: "100%",
        maxWidth: props.width || 400,
        height: open ? (props.height || 600) : 44,
        borderRadius: 12,
        overflow: "hidden",
        border: "1px solid #d1d5db",
        background: "#ffffff",
        display: "flex",
        flexDirection: "column"
      }}
    >
      <div
        style={{
          minHeight: 44,
          flexShrink: 0,
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          padding: "0 12px",
          background: "#1f2937",
          color: "#ffffff",
          fontWeight: 600,
          borderBottom: "1px solid #0f172a"
        }}
      >
        <span>{props.title || "Assistant"}</span>
        <button
          onClick={() => setOpen(v => !v)}
          style={{
            border: "none",
            background: "transparent",
            color: "#ffffff",
            cursor: "pointer",
            fontSize: 16,
            lineHeight: 1
          }}
          aria-label={open ? "Minimize chat" : "Expand chat"}
          title={open ? "Minimize" : "Expand"}
        >
          {open ? "-" : "+"}
        </button>
      </div>

      {open && (
        <div style={{ flex: 1, minHeight: 0, background: "#ffffff" }}>
          {useIframeAgent ? (
            <iframe
              src={trimmedAgentUrl}
              title={props.title || "Assistant"}
              style={{ height: "100%", border: "none", width: "100%" }}
              allow="clipboard-read; clipboard-write"
            />
          ) : props.tokenEndpointUrl && props.tokenEndpointUrl.trim() ? (
            <div ref={hostRef} style={{ height: "100%" }} />
          ) : (
            <div style={{ padding: 12, color: "#374151" }}>
              Set Agent URL or Token Endpoint URL in web part properties.
            </div>
          )}
        </div>
      )}
    </div>
  );
}
