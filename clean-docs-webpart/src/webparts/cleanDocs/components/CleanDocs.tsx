
import * as React from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { ICleanDocsProps } from "./ICleanDocsProps";
import styles from "./CleanDocs.module.scss";

// Calls the SharePoint REST API and returns raw item rows from the target document library.
//
// Expected behavior:
// - Only file items are returned (folders are excluded with FSObjType eq 0).
// - The query is capped by maxItems for predictable payload size.
// - If the request fails, the thrown error includes the response body to simplify diagnostics.
async function loadDocs(
  spHttpClient: SPHttpClient,
  webUrl: string,
  libraryTitle: string,
  maxItems: number
) {
  // Escape apostrophes in list titles (rare but possible), because list titles are embedded
  // in a single-quoted GetByTitle('...') segment and unescaped apostrophes break the URL.
  const safeTitle = libraryTitle.replace(/'/g, "''");

  // Build a REST endpoint that requests only fields we need for rendering.
  //
  // Field notes:
  // - FileRef: server-relative path used for navigation.
  // - FileLeafRef: physical file name fallback when DisplayText is blank.
  // - DisplayText: custom display label shown to users when provided.
  // - DisplayStartDate / DisplayEndDate: currently loaded for compatibility/future use.
  // - FSObjType: used in filter to ensure only files are returned.
  const url =
    `${webUrl}/_api/web/lists/GetByTitle('${safeTitle}')/items` +
    `?$select=FileRef,FileLeafRef,DisplayText,DisplayStartDate,DisplayEndDate,FSObjType` +
    `&$filter=FSObjType eq 0` +
    `&$top=${maxItems}`;

  // Execute the request through SPHttpClient (already scoped/authenticated by SPFx context).
  const res: SPHttpClientResponse = await spHttpClient.get(
    url,
    SPHttpClient.configurations.v1,
    {
      headers: {
        // OData metadata=none trims payload size by excluding verbose metadata blocks.
        "accept": "application/json;odata.metadata=none"
      }
    }
  );

  if (!res.ok) {
    // Read the body because SharePoint often includes the real failure reason there
    // (for example: missing field, misspelled library title, permission issue).
    const body = await res.text();
    throw new Error(`SharePoint REST error ${res.status}: ${body}`);
  }

  // Response shape is { value: [...] } for this endpoint when using metadata=none.
  const json = await res.json();
  return json.value;
}

// Represents a single file item returned from the document library query.
// Nullable fields are modeled as string | null because SharePoint commonly returns null
// for empty optional columns.
interface DocItem {
  FileRef: string;
  FileLeafRef: string;
  DisplayText?: string | null;
  DisplayStartDate: string | null;
  DisplayEndDate: string | null;
}

// Determines which label users see for a document.
// Preference order:
// 1) DisplayText (trimmed), if provided
// 2) FileLeafRef (actual file name), as fallback
function getDocDisplayName(doc: DocItem): string {
  const displayText = doc.DisplayText?.trim();
  return displayText ? displayText : doc.FileLeafRef;
}

export const CleanDocs: React.FC<ICleanDocsProps> = ({
  siteUrl,
  titleBarText,
  libraryTitle,
  maxItems,
  spHttpClient
}) => {
  // docs: final data set rendered as list items.
  const [docs, setDocs] = React.useState<DocItem[]>([]);
  // error: non-null when fetch fails; displayed in UI instead of the list.
  const [error, setError] = React.useState<string | null>(null);
  // loading: true until first fetch resolves/rejects.
  const [loading, setLoading] = React.useState(true);

  // Refetch when key inputs change.
  // Note: spHttpClient is stable from SPFx context and intentionally omitted here.
  React.useEffect(() => {
    loadDocs(spHttpClient, siteUrl, libraryTitle, maxItems)
      .then(items => {
        // Sort by the same label users see, so visual order aligns with displayed names.
        // Use case-insensitive comparison for friendlier alphabetical ordering.
        const sortedItems = [...items].sort((a, b) =>
          getDocDisplayName(a).localeCompare(getDocDisplayName(b), undefined, { sensitivity: "base" })
        );

        setDocs(sortedItems);
        // Fetch complete (success path).
        setLoading(false);
      })
      .catch(err => {
        // Surface the message to the user for easier troubleshooting.
        setError(err.message);
        // Fetch complete (error path).
        setLoading(false);
      });
  }, [siteUrl, libraryTitle, maxItems]);

  // Content area switches between loading, error, and loaded states.
  let content: React.ReactElement;

  if (loading) {
    content = React.createElement("div", null, "Loading...");
  } else if (error) {
    content = React.createElement("div", null, `Error: ${error}`);
  } else {
    content = React.createElement(
      "ul",
      null,
      docs.map(doc => {
        // FileRef may already be absolute depending on tenant/settings; normalize relative values.
        const docUrl = doc.FileRef.startsWith("http")
          ? doc.FileRef
          : new URL(doc.FileRef, window.location.origin).toString();

        return React.createElement(
          "li",
          { key: doc.FileRef },
          React.createElement(
            "a",
            {
              href: docUrl,
              // Always open docs in a new tab to keep the current SharePoint page intact.
              target: "_blank",
              rel: "noopener noreferrer",
              // Disable SharePoint client-side link interception so the browser follows the real URL.
              "data-interception": "off",
              onClick: (ev: React.MouseEvent<HTMLAnchorElement>) => {
                // Defensive: stop framework/page handlers from hijacking navigation behavior.
                ev.preventDefault();
                ev.stopPropagation();
                // Explicitly open new tab/window with safe window features.
                window.open(docUrl, "_blank", "noopener,noreferrer");
              }
            },
            getDocDisplayName(doc)
          )
        );
      })
    );
  }

  return React.createElement(
    "div",
    { className: styles.cleanDocs },
    // Optional configurable header text with safe default.
    React.createElement("div", { className: styles.titleBar }, titleBarText || "Clean Documents"),
    content
  );
};
