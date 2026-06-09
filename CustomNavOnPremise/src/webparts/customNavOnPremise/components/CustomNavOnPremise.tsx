import * as React from 'react';
import styles from './CustomNavOnPremise.module.scss';
import { ICustomNavOnPremiseProps } from './ICustomNavOnPremiseProps';

export interface ISiteLink {
  title: string;
  url: string;
}

export interface ICustomNavOnPremiseState {
  isLoading: boolean;
  errorMessage?: string;
  sites: ISiteLink[];
  customTree?: INavNode[];
}

interface INavNode {
  title: string;
  url: string;
  children: INavNode[];
}

interface ICustomNavigationItem {
  title: string;
  url: string;
  children?: ICustomNavigationItem[];
}

export default class CustomNavOnPremise extends React.Component<ICustomNavOnPremiseProps, ICustomNavOnPremiseState> {

  constructor(props: ICustomNavOnPremiseProps) {
    super(props);

    this.state = {
      isLoading: true,
      sites: []
    };
  }

  public componentDidMount(): void {
    const customTree: INavNode[] | undefined = this._parseCustomNavigation();

    if (customTree && customTree.length > 0) {
      this.setState({
        isLoading: false,
        errorMessage: undefined,
        sites: [],
        customTree: customTree
      });
      return;
    }

    this._loadAccessibleSites();
  }

  private _parseCustomNavigation(): INavNode[] | undefined {
    const rawValue: string = (this.props.customNavigationJson || '').trim();

    if (!rawValue) {
      return undefined;
    }

    try {
      const parsed: any = JSON.parse(rawValue);
      if (!Array.isArray(parsed)) {
        throw new Error('Custom navigation must be a JSON array.');
      }

      const toNode = (item: ICustomNavigationItem): INavNode => {
        const childItems: ICustomNavigationItem[] = Array.isArray(item.children) ? item.children : [];

        if (!item || !item.title || !item.url) {
          throw new Error('Each navigation item must include title and url.');
        }

        return {
          title: item.title,
          url: item.url.replace(/\/$/, ''),
          children: childItems.map((childItem: ICustomNavigationItem) => toNode(childItem))
        };
      };

      return parsed.map((item: ICustomNavigationItem) => toNode(item));
    } catch (error) {
      this.setState({
        isLoading: false,
        errorMessage: 'Invalid custom navigation JSON. Check web part properties.',
        sites: [],
        customTree: []
      });
      return [];
    }
  }

  private _loadAccessibleSites(): void {
    const searchApiUrl: string = this.props.siteUrl +
      "/_api/search/query?querytext='(contentclass:STS_Site OR contentclass:STS_Web)'&rowLimit=500&selectProperties='Title,Path'";

    fetch(searchApiUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=nometadata'
      },
      credentials: 'same-origin'
    })
      .then((response: Response) => {
        if (!response.ok) {
          throw new Error('Unable to load sites for the navigation bar.');
        }

        return response.json();
      })
      .then((searchResult: any) => {
        const rows: any[] = this._getRowsFromSearchResponse(searchResult);
        const seen: { [key: string]: boolean } = {};
        const sites: ISiteLink[] = [];

        rows.forEach((row: any) => {
          const cells: any[] = this._getCellsFromRow(row);
          const siteTitle: string = this._getCellValue(cells, 'Title');
          const sitePath: string = this._getCellValue(cells, 'Path');
          const normalizedPath: string = sitePath ? sitePath.replace(/\/$/, '') : '';

          if (!normalizedPath || seen[normalizedPath.toLowerCase()]) {
            return;
          }

          seen[normalizedPath.toLowerCase()] = true;
          sites.push({
            title: siteTitle || normalizedPath,
            url: normalizedPath
          });
        });

        this.setState({
          isLoading: false,
          errorMessage: undefined,
          sites: sites,
          customTree: undefined
        });
      })
      .catch((error: Error) => {
        this.setState({
          isLoading: false,
          errorMessage: error.message,
          sites: [],
          customTree: undefined
        });
      });
  }

  private _getRowsFromSearchResponse(searchResult: any): any[] {
    if (searchResult && searchResult.PrimaryQueryResult && searchResult.PrimaryQueryResult.RelevantResults &&
      searchResult.PrimaryQueryResult.RelevantResults.Table && searchResult.PrimaryQueryResult.RelevantResults.Table.Rows) {
      const rowsNoMetadata: any = searchResult.PrimaryQueryResult.RelevantResults.Table.Rows;
      return Array.isArray(rowsNoMetadata) ? rowsNoMetadata : (rowsNoMetadata.results || []);
    }

    if (searchResult && searchResult.d && searchResult.d.query && searchResult.d.query.PrimaryQueryResult &&
      searchResult.d.query.PrimaryQueryResult.RelevantResults && searchResult.d.query.PrimaryQueryResult.RelevantResults.Table &&
      searchResult.d.query.PrimaryQueryResult.RelevantResults.Table.Rows &&
      searchResult.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results) {
      return searchResult.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
    }

    return [];
  }

  private _getCellsFromRow(row: any): any[] {
    if (!row || !row.Cells) {
      return [];
    }

    return Array.isArray(row.Cells) ? row.Cells : (row.Cells.results || []);
  }

  private _getCellValue(cells: any[], key: string): string {
    const lowerKey: string = key.toLowerCase();

    for (let i: number = 0; i < cells.length; i++) {
      if (cells[i] && cells[i].Key && cells[i].Key.toLowerCase() === lowerKey) {
        return cells[i].Value || '';
      }
    }

    return '';
  }

  private _isCurrentSite(url: string): boolean {
    return this.props.currentWebUrl.replace(/\/$/, '').toLowerCase() === url.replace(/\/$/, '').toLowerCase();
  }

  private _getParentPath(url: string): string {
    const normalizedUrl: string = url.replace(/\/$/, '');
    const protocolIndex: number = normalizedUrl.indexOf('://');
    const pathStart: number = protocolIndex >= 0 ? normalizedUrl.indexOf('/', protocolIndex + 3) : normalizedUrl.indexOf('/');

    if (pathStart < 0) {
      return '';
    }

    const hostRoot: string = normalizedUrl.substring(0, pathStart);
    const pathPart: string = normalizedUrl.substring(pathStart);
    const pathSegments: string[] = pathPart.split('/').filter((segment: string) => segment.length > 0);

    if (pathSegments.length <= 1) {
      return '';
    }

    pathSegments.pop();
    return hostRoot + '/' + pathSegments.join('/');
  }

  private _buildNavigationTree(sites: ISiteLink[]): INavNode[] {
    const nodeMap: { [key: string]: INavNode } = {};
    const roots: INavNode[] = [];

    sites.forEach((site: ISiteLink) => {
      nodeMap[site.url.toLowerCase()] = {
        title: site.title,
        url: site.url,
        children: []
      };
    });

    const getNearestParent = (url: string): INavNode => {
      let parentPath: string = this._getParentPath(url).toLowerCase();

      while (parentPath) {
        if (nodeMap[parentPath]) {
          return nodeMap[parentPath];
        }

        parentPath = this._getParentPath(parentPath).toLowerCase();
      }

      return undefined;
    };

    sites.forEach((site: ISiteLink) => {
      const node: INavNode = nodeMap[site.url.toLowerCase()];
      const parent: INavNode = getNearestParent(site.url);

      if (parent) {
        parent.children.push(node);
      } else {
        roots.push(node);
      }
    });

    const sortTree = (nodes: INavNode[]): void => {
      nodes.sort((a: INavNode, b: INavNode) => a.title.localeCompare(b.title));
      nodes.forEach((node: INavNode) => sortTree(node.children));
    };

    sortTree(roots);
    return roots;
  }

  private _treeHasCurrentSite(node: INavNode): boolean {
    if (this._isCurrentSite(node.url)) {
      return true;
    }

    for (let i: number = 0; i < node.children.length; i++) {
      if (this._treeHasCurrentSite(node.children[i])) {
        return true;
      }
    }

    return false;
  }

  public render(): React.ReactElement<ICustomNavOnPremiseProps> {
    const navigationTree: INavNode[] = this.state.customTree || this._buildNavigationTree(this.state.sites);

    return (
      <div className={ styles.customNavOnPremise }>
        {this.state.isLoading && <div className={ styles.status }>Loading sites...</div>}
        {!this.state.isLoading && this.state.errorMessage && <div className={ styles.statusError }>{this.state.errorMessage}</div>}
        {!this.state.isLoading && !this.state.errorMessage && this.state.sites.length === 0 &&
          <div className={ styles.status }>No accessible sites found.</div>}

        {!this.state.isLoading && !this.state.errorMessage && navigationTree.length > 0 &&
          <ul className={ styles.navList } aria-label={ this.props.description || 'Site navigation' }>
            {navigationTree.map((site: INavNode) => {
              const isCurrent: boolean = this._treeHasCurrentSite(site);
              const hasChildren: boolean = site.children.length > 0;

              return (
                <li className={ hasChildren ? styles.navItemWithMenu : styles.navItem } key={ site.url }>
                  <a
                    className={
                      hasChildren
                        ? (isCurrent ? styles.navLinkWithMenuActive : styles.navLinkWithMenu)
                        : (isCurrent ? styles.navLinkActive : styles.navLink)
                    }
                    href={ site.url }
                    title={ site.title }
                  >
                    {site.title}
                  </a>

                  {hasChildren &&
                    <ul className={ styles.dropdownMenu }>
                      {site.children.map((childSite: INavNode) => {
                        const isChildCurrent: boolean = this._treeHasCurrentSite(childSite);

                        return (
                          <li className={ styles.dropdownItem } key={ childSite.url }>
                            <a
                              className={ isChildCurrent ? styles.dropdownLinkActive : styles.dropdownLink }
                              href={ childSite.url }
                              title={ childSite.title }
                            >
                              {childSite.title}
                            </a>
                          </li>
                        );
                      })}
                    </ul>
                  }
                </li>
              );
            })}
          </ul>
        }
      </div>
    );
  }
}
