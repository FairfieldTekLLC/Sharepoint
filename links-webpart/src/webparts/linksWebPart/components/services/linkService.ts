import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import type { WebPartContext } from "@microsoft/sp-webpart-base";

import { ILinkItem } from "./ILinkItem";

export async function getLinks(
  context: WebPartContext,
  listTitle: string,
  maxItems: number
): Promise<ILinkItem[]> {

  const sp = spfi().using(SPFx(context));

  // NOTE:
  // - If you DON'T have Active/Order columns, remove filter/orderBy.
  // - If your hyperlink column isn't named "Link", change select fields.
  const items = await sp.web.lists.getByTitle(listTitle).items
    .select("Id", "Title", "Link", "Category", "Description", "Order", "Active")
    .filter("Active eq 1")              // remove if you don't have Active
    .orderBy("Order", true)             // remove if you don't have Order
    .top(maxItems)();

  return items as unknown as ILinkItem[];
}
