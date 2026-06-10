
import * as React from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { ICleanDocsProps } from "./ICleanDocsProps";
import styles from "./CleanDocs.module.scss";

// Represents a single file item rendered in the document table.
interface DocItem {
  FileRef: string;
  FileLeafRef: string;
  DisplayText?: string;
  DisplayStartDate?: string;
  DisplayEndDate?: string;
  SortOrder?: number | string;
  DisplayDescription?: string;
}

const requiredColumns: string[] = ["SortOrder", "DisplayDescription"];

function getMissingColumnFromError(body: string): string | undefined {
  const match = /field or property '([^']+)' does not exist/i.exec(body);
  return match?.[1];
}

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
): Promise<DocItem[]> {
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
    `?$select=FileRef,FileLeafRef,DisplayText,DisplayStartDate,DisplayEndDate,SortOrder,DisplayDescription,FSObjType` +
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

    const missingColumn = getMissingColumnFromError(body);
    if (missingColumn && requiredColumns.indexOf(missingColumn) >= 0) {
      throw new Error(
        `Missing required column '${missingColumn}' in library '${libraryTitle}'. Add the column and try again.`
      );
    }

    throw new Error(`SharePoint REST error ${res.status}: ${body}`);
  }

  // Response shape is { value: [...] } for this endpoint when using metadata=none.
  const json = await res.json() as { value: Array<Record<string, unknown>> };
  return json.value.map(item => ({
    FileRef: String(item.FileRef ?? ""),
    FileLeafRef: String(item.FileLeafRef ?? ""),
    DisplayText: typeof item.DisplayText === "string" ? item.DisplayText : undefined,
    DisplayStartDate: typeof item.DisplayStartDate === "string" ? item.DisplayStartDate : undefined,
    DisplayEndDate: typeof item.DisplayEndDate === "string" ? item.DisplayEndDate : undefined,
    SortOrder: typeof item.SortOrder === "number" || typeof item.SortOrder === "string" ? item.SortOrder : undefined,
    DisplayDescription: typeof item.DisplayDescription === "string" ? item.DisplayDescription : undefined
  }));
}

// Determines which label users see for a document.
// Preference order:
// 1) DisplayText (trimmed), if provided
// 2) FileLeafRef (actual file name), as fallback
function getDocDisplayName(doc: DocItem): string {
  const displayText = doc.DisplayText?.trim();
  return displayText ? displayText : doc.FileLeafRef;
}

function getSortOrderValue(doc: DocItem): number {
  const rawValue = doc.SortOrder;

  if (typeof rawValue === "number" && Number.isFinite(rawValue)) {
    return rawValue;
  }

  if (typeof rawValue === "string") {
    const parsed = Number(rawValue.trim());
    if (!Number.isNaN(parsed)) {
      return parsed;
    }
  }

  return Number.MAX_SAFE_INTEGER;
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
        // Sort by explicit SortOrder first, then by visible name for deterministic ordering.
        const sortedItems = [...items].sort((a, b) =>
          getSortOrderValue(a) - getSortOrderValue(b) ||
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

        const description = doc.DisplayDescription?.trim() || "";

        return React.createElement(
          "li",
          { key: doc.FileRef, className: styles.docItem },
          React.createElement(
            React.Fragment,
            null,
            React.createElement(
              "a",
              {
                className: styles.docLink,
                href: docUrl,
                title: description,
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
            ),
            description
              ? React.createElement(
                "div",
                { className: styles.docDescription },
                description
              )
              : null
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
