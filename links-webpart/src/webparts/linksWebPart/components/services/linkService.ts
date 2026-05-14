/**
 * linkService.ts  (links-webpart)
 *
 * Data-access service for the LinksWebPart component.
 * Provides a single exported async function, `getLinks`, that queries a
 * SharePoint list via the PnP SP library and returns typed ILinkItem objects.
 *
 * Prerequisites:
 *  - The target list must expose the columns: Id, Title, Link, Category,
 *    Description, Order, Active.
 *  - The 'Active' column should be a Yes/No (boolean) column.
 *    Remove the `filter` clause if the column does not exist on the list.
 *  - The 'Order' column should be a numeric column.
 *    Remove the `orderBy` clause if the column does not exist on the list.
 *  - The 'Link' column must be a SharePoint Hyperlink column.
 */
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import type { WebPartContext } from "@microsoft/sp-webpart-base";

import { ILinkItem } from "./ILinkItem";

/**
 * Fetches link items from the named SharePoint list.
 *
 * Uses the PnP SP fluent API (spfi + SPFx) for authenticated REST calls.
 * Items are filtered to Active = 1 and ordered by the Order column ascending.
 *
 * NOTE: If your list does not have Active or Order columns, remove the
 * corresponding `.filter()` or `.orderBy()` calls to avoid REST query errors.
 *
 * @param context   - The SPFx WebPartContext; passed to SPFx() so PnP SP can
 *                    authenticate using the current user's SharePoint session.
 * @param listTitle - Display name of the SharePoint list to query (case-insensitive).
 * @param maxItems  - Maximum number of items to retrieve (OData `$top` value).
 * @returns         - A promise resolving to an array of ILinkItem objects.
 */
export async function getLinks(
  context: WebPartContext,
  listTitle: string,
  maxItems: number
): Promise<ILinkItem[]> {

  // Initialise the PnP SP client, scoped to the current web via the SPFx context.
  // `spfi().using(SPFx(context))` injects the correct authentication cookie/token
  // and sets the web URL to the current SharePoint site automatically.
  const sp = spfi().using(SPFx(context));

  // NOTE:
  // - If you DON'T have Active/Order columns, remove filter/orderBy.
  // - If your hyperlink column isn't named "Link", change select fields.
  const items = await sp.web.lists.getByTitle(listTitle).items
    // Select only the columns required for rendering to minimise payload size.
    .select("Id", "Title", "Link", "Category", "Description", "Order", "Active")
    // Only retrieve items where the Active flag is set to 1 (true).
    .filter("Active eq 1")              // remove if you don't have Active
    // Sort ascending by the numeric Order column so display order matches editorial intent.
    .orderBy("Order", true)             // remove if you don't have Order
    // Cap the result set at the configured maximum to prevent over-fetching.
    .top(maxItems)();

  // The PnP SP response is typed as `any[]`; cast to ILinkItem[] via `unknown`
  // because the selected fields align with the interface definition.
  return items as unknown as ILinkItem[];
}
