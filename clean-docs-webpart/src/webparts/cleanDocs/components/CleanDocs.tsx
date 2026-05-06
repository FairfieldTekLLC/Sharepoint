
import * as React from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { ICleanDocsProps } from "./ICleanDocsProps";
import styles from "./CleanDocs.module.scss";

async function loadDocs(
  spHttpClient: SPHttpClient,
  webUrl: string,
  libraryTitle: string,
  maxItems: number
) {
  // Escape apostrophes in list titles (rare but can happen)
  const safeTitle = libraryTitle.replace(/'/g, "''");

  const url =
    `${webUrl}/_api/web/lists/GetByTitle('${safeTitle}')/items` +
    `?$select=FileRef,FileLeafRef,DisplayStartDate,DisplayEndDate,FSObjType` +
    `&$filter=FSObjType eq 0` +
    `&$top=${maxItems}`;

  const res: SPHttpClientResponse = await spHttpClient.get(
    url,
    SPHttpClient.configurations.v1,
    {
      headers: {
        // OData v4 header (recommended)
        "accept": "application/json;odata.metadata=none"
      }
    }
  );

  if (!res.ok) {
    // IMPORTANT: read error body for the true reason (missing field, wrong list title, etc.)
    const body = await res.text();
    throw new Error(`SharePoint REST error ${res.status}: ${body}`);
  }

  const json = await res.json();
  return json.value;
}

interface DocItem {
  FileRef: string;
  FileLeafRef: string;
  DisplayStartDate: string | null;
  DisplayEndDate: string | null;
}

export const CleanDocs: React.FC<ICleanDocsProps> = ({ siteUrl, libraryTitle, maxItems, spHttpClient }) => {
  const [docs, setDocs] = React.useState<DocItem[]>([]);
  const [error, setError] = React.useState<string | null>(null);
  const [loading, setLoading] = React.useState(true);

  React.useEffect(() => {
    loadDocs(spHttpClient, siteUrl, libraryTitle, maxItems)
      .then(items => {
        setDocs(items);
        setLoading(false);
      })
      .catch(err => {
        setError(err.message);
        setLoading(false);
      });
  }, [siteUrl, libraryTitle, maxItems]);

  if (loading) return React.createElement("div", null, "Loading...");
  if (error) return React.createElement("div", null, `Error: ${error}`);

  return React.createElement(
    "div",
    { className: styles.cleanDocs },
    React.createElement(
      "ul",
      null,
      docs.map(doc =>
      React.createElement(
        "li",
        { key: doc.FileRef },
        React.createElement("a", { href: doc.FileRef }, doc.FileLeafRef)
      )
    )
    )
  );
};
