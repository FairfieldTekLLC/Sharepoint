# SharePoint SPFx Solutions

This repository contains five SharePoint Framework solutions:

1. `ChatBotWebPart`
2. `clean-docs-webpart`
3. `CustomNav`
4. `links-webpart`
5. `spfx-links-webpart`

All five projects use SPFx `1.22.2`, React `17`, and Heft-based build scripts. They are packaged independently and can be deployed separately.

## Repository layout

| Project | Type | Purpose | Output package |
| --- | --- | --- | --- |
| `ChatBotWebPart` | Client-side web part | Embeds a chatbot either through an iframe URL or Bot Framework Web Chat + Direct Line token flow | `ChatBotWebPart/sharepoint/solution/chat-bot-web-part.sppkg` |
| `clean-docs-webpart` | Client-side web part | Displays a cleaned-up document library listing with friendly display text and forced new-tab open behavior | `clean-docs-webpart/sharepoint/solution/clean-docs-webpart.sppkg` |
| `CustomNav` | Application Customizer | Injects a top navigation bar built from Microsoft Graph site search results and optional external links | `CustomNav/sharepoint/solution/custom-nav.sppkg` |
| `links-webpart` | Client-side web part | Displays categorized links from a SharePoint list using PnP/SP and configurable styling | `links-webpart/sharepoint/solution/links-webpart.sppkg` |
| `spfx-links-webpart` | Client-side web part | Displays links from a selected SharePoint list using direct REST calls and optional category/date filtering | `spfx-links-webpart/sharepoint/solution/spfx-links-webpart.sppkg` |

## Development prerequisites

All projects share the same baseline tooling requirements:

- Node.js `>=22.14.0 <23.0.0`
- npm
- SharePoint Online app catalog access for deployment
- A Microsoft 365 tenant with SPFx enabled

Optional but useful:

- Azure CLI if you want to use the `CustomNav/scripts/Register-CustomNav.ps1` helper
- A SharePoint test site collection for validating deployment and data configuration

## Common local setup

Each project is standalone. Install dependencies inside the solution folder you want to work on.

```powershell
cd c:\SourceTree\Sharepoint\ChatBotWebPart
npm install
npm start
```

The same pattern applies to each project folder:

- `npm install` installs dependencies
- `npm start` runs `heft start --clean`
- `npm run build` runs a production build and packages the `.sppkg`
- `npm run clean` removes build output

## Build and packaging

For any solution:

```powershell
cd c:\SourceTree\Sharepoint\<solution-folder>
npm install
npm run build
```

That build script runs:

```text
heft test --clean --production && heft package-solution --production
```

After the build completes, upload the generated `.sppkg` from the project's `sharepoint/solution` folder to your SharePoint App Catalog.

## Deployment model

All five projects are configured with:

- `includeClientSideAssets: true`
- `skipFeatureDeployment: true`
- `isDomainIsolated: false`

That means:

- Client-side assets are packaged into the SPFx solution package.
- You can choose tenant-wide deployment from the app catalog.
- Web parts still need to be added to a page before users see them.
- The `CustomNav` extension may still require site installation or explicit custom action registration depending on how you deploy it.

## Solution details

## ChatBotWebPart

### What it does

`ChatBotWebPart` renders a chat widget with two supported modes:

1. `iframe` mode: if `agentUrl` is populated, the web part embeds the published chatbot URL inside an iframe.
2. `Web Chat` mode: if `agentUrl` is blank and `tokenEndpointUrl` is set, the component loads the Bot Framework Web Chat client and requests a Direct Line token from the configured endpoint.

In Web Chat mode, the component also tries to perform token exchange through `aadTokenProviderFactory` so OAuth cards can complete silently for the signed-in SharePoint user.

### Key properties

Configured from the web part property pane:

- `agentUrl`: published Copilot Studio or Azure bot iframe URL
- `tokenEndpointUrl`: Direct Line token endpoint used when `agentUrl` is blank
- `title`: header text shown in the widget
- `height`: widget height in pixels
- `width`: widget width in pixels

### Setup notes

- No SharePoint list or library is required.
- If you use iframe mode, make sure the target bot URL allows embedding in SharePoint.
- If you use Web Chat mode, the token endpoint must return a Direct Line token payload that the component can parse.
- The component dynamically loads the Web Chat script from the Bot Framework CDN when needed.

### Deployment notes

- Build and upload `chat-bot-web-part.sppkg`.
- Add the web part to a SharePoint page.
- Open the property pane and configure either:
	- `agentUrl`, or
	- `tokenEndpointUrl`
- If both are blank, the web part renders a configuration warning.

### Important information

- The web part supports `SharePointWebPart`, `TeamsPersonalApp`, `TeamsTab`, and `SharePointFullPage` hosts.
- The current manifest still contains placeholder title/description text. The runtime configuration is what actually controls behavior.

## clean-docs-webpart

### What it does

`clean-docs-webpart` queries a SharePoint document library through the SharePoint REST API and renders a simplified list of files. It sorts the output by the display label users actually see, and all links open in a new browser tab.

The component prefers `DisplayText` when present, otherwise it falls back to `FileLeafRef`.

### Key properties

- `titleBarText`: title shown above the list
- `libraryTitle`: SharePoint library title to query
- `maxItems`: maximum number of file rows to request

### Expected library configuration

The library should expose these fields because the component selects them directly:

| Field | Purpose | Required |
| --- | --- | --- |
| `FileRef` | Builds the document URL | Yes |
| `FileLeafRef` | File name fallback | Yes |
| `DisplayText` | Friendly label shown to users | Recommended |
| `DisplayStartDate` | Loaded by the component but not currently enforced | Optional |
| `DisplayEndDate` | Loaded by the component but not currently enforced | Optional |
| `FSObjType` | Used to exclude folders | Yes |

### Setup steps

1. Create or identify a document library.
2. Add a `DisplayText` column if you want friendly names instead of file names.
3. Add the web part to a page.
4. Set `libraryTitle` to the exact library title.
5. Adjust `maxItems` and `titleBarText` as needed.

### Deployment notes

- Build and upload `clean-docs-webpart.sppkg`.
- Install the app to the target site if needed.
- Add the web part to a modern page.

### Important information

- The component queries by library title, so renaming the library means you must update the property pane value.
- Failed REST requests bubble up into a user-visible error message, including the response body when available.

## CustomNav

### What it does

`CustomNav` is an SPFx Application Customizer, not a web part. It injects a navigation bar into the page placeholder area and builds menu items from Microsoft Graph site search results.

Behavior summary:

- Prefers the `Top` placeholder
- Falls back to the `Bottom` placeholder
- Falls back again to a fixed host attached to `document.body` if neither placeholder exists
- Caches site results in `localStorage` for 2 minutes
- Filters personal sites and optionally team-site paths
- Supports additional external links from extension properties

### How it decides what to show

The extension loads accessible sites through `SiteSearchService`, which calls `POST /search/query` in Microsoft Graph for `entityTypes: ['site']`. It then:

1. normalizes site URLs
2. filters unwanted personal/team paths
3. builds a parent-child tree from site paths
4. renders dropdown and flyout menus

### Configuration model

The strongly-typed main property interface is empty, but the code reads these optional properties when provided as custom action JSON:

| Property | Purpose | Default behavior |
| --- | --- | --- |
| `hidePersonalSites` | Hides OneDrive/personal sites | `true` |
| `hideTeamSites` | Hides team sites based on path prefixes | `true` |
| `teamSitePathPrefixes` | Prefixes such as `/teams/` used for exclusion | `['/teams/']` |
| `externalLinks` | Additional static links appended to the nav | none |

`externalLinks` should be an array of objects shaped like this:

```json
[
	{
		"title": "Intranet Home",
		"url": "https://tenant.sharepoint.com/sites/intranet",
		"target": "_self"
	},
	{
		"title": "External Resource",
		"url": "https://example.com",
		"target": "_blank"
	}
]
```

### Deployment steps

1. Build and upload `custom-nav.sppkg`.
2. Deploy it in the app catalog.
3. Install it on the target site if the extension does not appear automatically.
4. Confirm the feature/custom action is active.

The solution contains `sharepoint/assets/elements.xml`, which registers a `ClientSideExtension.ApplicationCustomizer` custom action for component ID `3b6f2fb3-8cd0-4f52-9d2e-0d2a4f24c8d1`.

### Alternate registration method

You can also register the application customizer directly against a site with:

`CustomNav/scripts/Register-CustomNav.ps1`

Usage:

```powershell
cd c:\SourceTree\Sharepoint\CustomNav\scripts
az login
.\Register-CustomNav.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite"
```

### Important information

- Packaging the solution successfully does not guarantee the nav appears. For app customizers, the app must be installed or the custom action must exist on the site.
- `webApiPermissionRequests` is currently empty in `config/package-solution.json`, even though the code uses Microsoft Graph search. If your tenant blocks the Graph call, you may need to add and consent to the appropriate Graph permissions before the menu can load.
- The extension renders a fallback message instead of failing hard when Graph is unavailable.

## links-webpart

### What it does

`links-webpart` displays a categorized list of links from a SharePoint list. It groups records by category, shows collapsible sections, and supports optional descriptions plus configurable color theming.

This project uses `@pnp/sp` and `@pnp/spfx-property-controls`.

### Key properties

- `listTitle`: list title to query
- `maxItems`: maximum links to load
- `openInNewTab`: open links in `_blank`
- `showTopText`: show the built-in hotline banner
- `showDescription`: show each item's description
- `backgroundColor`: card background color
- `titleBarColor`: section header color
- `linkTextColor`: link text color

### Expected list configuration

The component selects and uses these fields directly:

| Field | Type | Purpose | Required |
| --- | --- | --- | --- |
| `Title` | Single line of text | Link text | Yes |
| `Link` | Hyperlink | Target URL | Yes |
| `Category` | Text or managed metadata | Grouping section | Recommended |
| `Description` | Multiple lines of text | Optional subtext | Optional |
| `Order` | Number | Sort order | Recommended |
| `Active` | Yes/No | Filters visible links | Recommended |

### Important list behavior

- The service applies `.filter("Active eq 1")`.
- The service applies `.orderBy("Order", true)`.
- If your list does not contain `Active` or `Order`, the current implementation must be changed before production use, or REST queries will fail.
- `Category` can be plain text, managed metadata, or multi-value data. The component normalizes those shapes into a readable label.

### Setup steps

1. Create a SharePoint list.
2. Add the columns shown above.
3. Populate the list with links.
4. Add the web part to a page.
5. Set `listTitle` to the exact list title.
6. Configure optional colors and display settings.

### Deployment notes

- Build and upload `links-webpart.sppkg`.
- Add the web part to a page after installing the solution.
- If data fails to load, verify the list title and required field internal names first.

## spfx-links-webpart

### What it does

`spfx-links-webpart` is another links solution, but it is implemented differently from `links-webpart`:

- it selects a list by GUID instead of title
- it uses `SPHttpClient` directly instead of PnP/SP
- it supports optional category filtering at query time
- it filters items client-side by start and stop dates

### Key properties

- `title`: heading above the list
- `listId`: SharePoint list GUID chosen from a property pane dropdown
- `maxItems`: max rows returned
- `showDescription`: show description text
- `openInNewTab`: link target behavior
- `category`: optional exact category filter

### Expected list configuration

The REST query expects these internal names:

| Field | Type | Purpose | Required |
| --- | --- | --- | --- |
| `Title` | Single line of text | Anchor text | Yes |
| `LinkUrl` | Hyperlink | Target URL | Yes |
| `LinkDescription` | Text | Description below the link | Optional |
| `Category` | Text | Optional server-side filter | Optional |
| `SortOrder` | Number | Sort order | Recommended |
| `StartDate` | Date/Time | Hide links before this date | Optional |
| `StopDate` | Date/Time | Hide links after this date | Optional |

### Important list behavior

- Items are fetched from `/lists(guid'...')/items`.
- Results are ordered by `SortOrder asc, Title asc`.
- If `category` is set, the web part adds an OData filter: `Category eq 'value'`.
- Links outside the `StartDate` / `StopDate` window are hidden client-side.
- The property pane dropdown only lists non-hidden SharePoint lists.

### Setup steps

1. Create a SharePoint list with the required columns.
2. Add the web part to a page.
3. Open the property pane.
4. Select the list from the dropdown.
5. Optionally set a category filter.

### Deployment notes

- Build and upload `spfx-links-webpart.sppkg`.
- Add the web part to a page after installation.
- If the list dropdown is empty, verify the current site actually contains non-hidden lists and that the solution has loaded correctly.

## Recommended deployment checklist

For any of these solutions:

1. Run `npm install` in the solution folder.
2. Run `npm run build`.
3. Upload the generated `.sppkg` to the tenant app catalog.
4. Choose whether to make it tenant-wide available.
5. Install the app on the target site if required.
6. Add the web part to a page, or verify the `CustomNav` custom action is active.
7. Configure the expected SharePoint list/library data source.
8. Validate with a non-admin user.

## Troubleshooting notes

- If `CustomNav` packages successfully but does not appear, verify that the app is installed on the site or that the custom action was actually created.
- If `links-webpart` fails, verify `listTitle`, `Active`, and `Order` field availability first.
- If `spfx-links-webpart` fails, verify the list contains `LinkUrl` and that the selected list GUID is correct.
- If `clean-docs-webpart` fails, verify the library title and confirm the library exposes the selected fields.
- If `ChatBotWebPart` fails in Web Chat mode, verify the token endpoint returns a valid Direct Line token and that any downstream auth flow allows the SharePoint user to complete token exchange.

## Project versions

| Project | package.json version | solution version |
| --- | --- | --- |
| `ChatBotWebPart` | `0.0.1` | `1.0.0.0` |
| `clean-docs-webpart` | `0.0.1` | `1.0.0.0` |
| `CustomNav` | `0.0.5` | `1.0.0.8` |
| `links-webpart` | `0.0.1` | `1.0.0.0` |
| `spfx-links-webpart` | `0.0.1` | `1.0.0.0` |

## Notes for future maintenance

- The per-project `README.md` files are still template placeholders. This root README is now the authoritative repository-level guide.
- If you rename any SharePoint fields, update the code and this document together.
- If `CustomNav` needs Graph access in a stricter tenant, add explicit `webApiPermissionRequests` and document the required admin consent.