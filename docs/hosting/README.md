# Hosting the Division 296 calculator

The site (`web/`) is a static, client-side app. The standalone build
(`web/standalone/div296-reset-calculator.html`) is one self-contained file with
no server dependency, which makes hosting flexible. This folder covers how to
get it in front of staff **inside Microsoft 365** (SharePoint / OneDrive / Teams)
when GitHub Pages / external hosting isn't wanted.

## The core constraint

SharePoint, OneDrive and Teams are **document stores, not web hosts**. Modern
SharePoint Online blocks custom scripts by default ("NoScript"), so an `.html`
file uploaded there **won't execute its JavaScript** when opened in the browser —
it downloads or shows inert. Teams *Files* live in the same SharePoint storage,
so they inherit the limit. A Teams **Website tab** can *display* a URL but does
not *host* one.

## Options, ranked

| # | Option | Live URL? | Needs | Best for |
|---|--------|-----------|-------|----------|
| 1 | **SPFx web part** | ✅ runs in SharePoint page + Teams tab | one-time App Catalog approval (admin) | the proper M365-native answer |
| 2 | **Azure Static Web Apps** + Teams Website tab | ✅ real internal URL, Entra-locked | an Azure subscription | a clean link without app-catalog approval |
| 3 | **Share the standalone file** (Teams/SharePoint/OneDrive) | ❌ download & open locally | nothing | works today, zero admin |

### 1. SPFx web part (recommended if IT will approve)

SharePoint Framework is Microsoft's supported way to run custom JS on a modern
page **without** weakening the NoScript setting. We package the calculator as a
web part, an admin approves it once in the App Catalog, then it can be added to
a SharePoint page and/or pinned as a Teams tab — visible only to staff with site
access (stays internal/proprietary).

- Source + build steps: [`spfx/BUILD.md`](spfx/BUILD.md)
- The note to hand IT: [`IT-request.md`](IT-request.md)

### 2. Azure Static Web Apps (recommended if you have Azure)

Free tier, gives a real URL (e.g. `https://div296.yourfirm.com`), and can be
**locked to your Entra ID tenant** so only staff can open it. Then drop that URL
into a Teams **Website tab**. Best if app-catalog approval isn't available but an
Azure subscription is. (Ask me and I'll add the `staticwebapp.config.json` +
deploy steps.)

### 3. Share the standalone file (works today, no admin)

Put `web/standalone/div296-reset-calculator.html` in a Teams channel's Files, a
SharePoint document library, or a shared drive. Staff click **Download** and
open it — it runs fully offline in their browser. Not a "website", but immediate
and requires nothing from IT.

## Decision shortcut

- **IT will approve an app** → Option 1 (SPFx). See `IT-request.md` + `spfx/BUILD.md`.
- **You have Azure** → Option 2.
- **Need it in front of people right now** → Option 3, and pursue 1 or 2 in parallel.

## Note on the proprietary concern

Options 1 and 2 keep the tool **internal** automatically (behind M365 / Entra
sign-in), which is why they're preferred over public GitHub Pages for a
proprietary tool.
