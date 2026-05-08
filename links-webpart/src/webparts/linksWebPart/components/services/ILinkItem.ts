export interface ILinkFieldValue {
  Url: string;
  Description?: string;
}

export interface ILinkItem {
  Id: number;
  Title: string;
  Link: ILinkFieldValue;
  Category?: unknown;
  Description?: string;
  Order?: number;
  Active?: boolean;
}
