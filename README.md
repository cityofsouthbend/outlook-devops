# Azure-Outlook DevOps Add-in

Outlook add-in for the City of South Bend that integrates with Azure DevOps. Create work items from emails, search/browse existing items, and insert rich work item cards into email drafts.

## Features

- **Create Ticket** (read mode only) - Create Issues, Requests, Projects, Tasks, or Sticky Notes directly from a selected email. Captures the full email body (with inline images resolved to DevOps-hosted URLs), file attachments, parent link, assignee, and current iteration.
- **Search Items** (read + compose) - Search work items by title or ID with optional type filtering. Results display as cards with type badge, state indicator, and assigned user. From results you can "Use as Parent" (switches to Create tab) or "Insert" (switches to Insert tab).
- **Insert Card** (compose mode only) - Insert a formatted work item card into an email draft at the cursor position. Cards use inline CSS for email client compatibility (no classes, table-based layout).

## Prerequisites

- Microsoft Outlook (desktop or web)
- Azure DevOps organization with a Personal Access Token (PAT)
- Node.js 18+

## Quick Start

```bash
git clone <repo-url>
cd outlook-devops
cp .env.example .env    # then edit with your PAT, org, and project
npm install
npm run build:dev
npm start               # sideloads into Outlook and starts dev server
```

## Configuration

### Environment Variables

Create a `.env` file in the project root (see `.env.example`):

| Variable | Description | Default |
|---|---|---|
| `DEVOPS_PAT` | Azure DevOps Personal Access Token | (none - required) |
| `DEVOPS_ORG` | Organization name | `southbendin` |
| `DEVOPS_PROJECT` | Project name, URL-encoded | `Digital%20-%20Product%20Portfolio` |

The PAT is bundled into the client-side JS at build time via `dotenv-webpack` (with `systemvars: true` so pipeline env vars also work). Auth is `"Basic " + btoa(":" + PAT)` - the username portion is empty.

### Team Members

Edit `src/taskpane/config/team-members.json` to update the assignment dropdown. Each entry has `{ "label": "Display Name", "value": "user@email" }`.

### Work Item Types

Edit `src/taskpane/config/work-item-types.json` to add or remove work item types shown in the Create and Search forms. Each entry has `{ "label": "Display Name", "value": "TypeName" }`.

## Architecture

### File Structure

```
src/
  taskpane/
    taskpane.js              # Entry point / orchestrator (~44 lines)
    taskpane.html            # Tabbed UI shell (Create | Search | Insert)
    taskpane.css             # All styles (copied to dist, linked from HTML)
    config/
      team-members.json      # Assignee dropdown options
      work-item-types.json   # Work item type options
    services/
      devops-api.js          # Azure DevOps REST API (WIQL, work items, attachments)
      outlook-mail.js        # Office.js wrappers (mail body, attachments, compose insert)
    components/
      tab-manager.js         # Tab switching + mode-based enable/disable
      create-ticket.js       # Create work item form + submission logic
      search-items.js        # Search UI + result card rendering
      insert-card.js         # Rich HTML card builder + compose insertion
      ui-helpers.js          # Shared DOM utilities (dropdowns, status banner, debounce)
  commands/
    commands.js              # Ribbon command handlers
    commands.html            # Command function file
assets/
  icon-*.png                 # Add-in icons (16, 32, 64, 80, 128)
manifest.xml                 # Office add-in manifest (dev URLs)
webpack.config.js            # Webpack 5 config
azure-pipelines.yml          # CI/CD pipeline definition
```

### Key Design Patterns

**Mode detection** - The add-in detects read vs. compose mode by checking `typeof Office.context.mailbox.item.subject`: string means read mode, object (setter) means compose mode. Tabs are enabled/disabled accordingly.

**Inter-module communication** - Modules are loosely coupled via custom DOM events:
- `parent-selected` - fired by search-items when "Use as Parent" is clicked; listened by create-ticket
- `item-selected-for-insert` - fired by search-items when "Insert" is clicked; listened by insert-card

**Initialization flow** (`taskpane.js`):
1. `Office.onReady` fires
2. Config-driven UI is populated (work item types, team dropdown, search filter)
3. `initTabs(mode)` sets up navigation and disables mode-inappropriate tabs
4. Module `init()` functions are called based on mode

**CSS loading** - CSS is copied to `dist/` via `CopyWebpackPlugin` and linked from `taskpane.html` with a `<link>` tag (not processed through CSS loaders).

### Inline Image Handling

When creating a ticket, the add-in:
1. Fetches the email HTML body via `body.getAsync`
2. Iterates over `item.attachments` where `isInline === true`
3. Calls `getAttachmentContentAsync` (requires Mailbox API 1.8) with a 15-second timeout
4. Uploads each image to DevOps as an attachment
5. Replaces `cid:` references in the HTML body with the DevOps attachment URLs
6. Sets the resolved HTML as the work item description field

Calls to `getAttachmentContentAsync` include a 200ms delay between each to prevent Office.js message channel conflicts. The content format is checked (`base64` vs `url`) since Office.js can return either.

### DevOps API Layer (`devops-api.js`)

All API calls go through three helpers: `apiGet`, `apiPost`, `apiPatch`. Each checks for JSON response (non-JSON means invalid/expired PAT). Key functions:

| Function | Purpose |
|---|---|
| `createWorkItem(type, patchOps)` | POST a new work item with `bypassrules=true` |
| `uploadAttachment(fileName, blob)` | POST binary blob to attachment endpoint |
| `addAttachmentToWorkItem(id, url, html, field)` | PATCH to attach file + set body field |
| `setCurrentIteration(id)` | GET current iteration, then PATCH the work item |
| `searchWorkItems(text, typeFilter)` | WIQL query by title/ID, then batch fetch details |
| `searchParentItems(text)` | Same as search but filtered to allowed parent types |
| `getGraphUsers()` | GET users from VSSPS Graph API (for CreatedBy resolution) |

Parent types allowed: Request, Issue, Project, Requirement, Detail, Planned Work, Maintenance, Product.

The body field is `Microsoft.VSTS.TCM.ReproSteps` for Issues, `System.Description` for everything else.

### Office.js API Layer (`outlook-mail.js`)

| Function | Purpose |
|---|---|
| `getMode()` | Returns `"read"`, `"compose"`, or `"unknown"` |
| `getReadMessageInfo()` | Returns `{ subject, from, attachments, itemId }` |
| `getEmailBody()` | Promise wrapper for `body.getAsync(Html)` |
| `getAttachmentContent(id, timeout)` | Promise wrapper for `getAttachmentContentAsync` with 15s timeout |
| `insertHtmlIntoCompose(html)` | Promise wrapper for `body.setSelectedDataAsync(Html)` |

## Build & Deployment

### Build

| Command | Description |
|---|---|
| `npm run build` | Production build (minified, prod URLs in manifest) |
| `npm run build:dev` | Development build (source maps, localhost URLs) |
| `npm run dev-server` | Start webpack dev server with HTTPS on port 3000 |
| `npm start` | Sideload manifest into Outlook and start dev server |
| `npm run validate` | Validate manifest.xml against Office schema |
| `npm run watch` | Development build with file watching |

Webpack replaces `https://localhost:3000/` with `https://outlookazureaddin.z22.web.core.windows.net/` in the manifest during production builds.

### CI/CD Pipeline

`azure-pipelines.yml` triggers on push to `main`:
1. Installs Node.js 18
2. Runs `npm ci && npm run build` with env vars (`DEVOPS_PAT`, `DEVOPS_ORG`, `DEVOPS_PROJECT`)
3. Deploys `dist/` to Azure Static Web Apps using `deployment_token`

**Pipeline variables** (set in Azure DevOps pipeline settings):
- `DEVOPS_PAT` - PAT token (bundled at build time)
- `DEVOPS_ORG` - Organization name
- `DEVOPS_PROJECT` - Project name (URL-encoded)
- `deployment_token` - Azure Static Web Apps deployment token

### Manifest

- `manifest.xml` - Source manifest with localhost URLs (for development)
- `dist/manifest.prod.xml` - Built manifest with production URLs (for sideloading/deployment)
- Requires `Mailbox` API version **1.8** (for `getAttachmentContentAsync`)
- Permissions: `ReadWriteMailbox`

### Hosting

Production site: `https://outlookazureaddin.z22.web.core.windows.net`

## Troubleshooting

- **"Expected JSON but received HTML"** - Your PAT token is missing, invalid, or expired. Regenerate it in Azure DevOps and rebuild.
- **"Failed to create ticket"** - Check that the PAT has Work Items (Read & Write) scope.
- **Add-in doesn't load** - Verify manifest URLs match your deployment target. For local dev, URLs should be `https://localhost:3000/`.
- **Insert tab disabled** - The Insert tab only appears in compose mode (new/reply/forward).
- **Create tab disabled** - The Create tab only appears in read mode (viewing a received email).
- **Inline images not appearing in DevOps** - Requires Mailbox API 1.8. Check the browser console for timeout or message channel errors. The add-in will gracefully skip images that fail to download.
- **"message channel closed" errors** - This is an Office.js runtime issue when the Outlook host can't keep up with async calls. The 200ms delay between attachment calls should mitigate this; if it persists, the timeout will catch it and skip the problematic attachment.

## .gitignore Notes

The `.gitignore` excludes `dist/**` and `config/**`. The `config/**` rule refers to a root-level config directory (not `src/taskpane/config/`). The JSON config files under `src/taskpane/config/` are tracked in git normally.
