# Request to IT — host the Division 296 calculator in SharePoint / Teams

**From:** Aiden Hiew · **What I'm asking for:** approval to add a small in-house
web part to our SharePoint App Catalog so an internal calculator can run on a
SharePoint page and as a Teams tab.

---

## What it is

A single-page **Division 296 reset calculator** — an internal tool that shows,
side by side, an SMSF's Division 296 tax if it resets its cost base vs if it
doesn't. It's the web version of an Excel model the firm already uses.

It is **100% client-side**: plain HTML, CSS and JavaScript. There is **no
server, no database, and no external network calls** — every calculation runs
in the user's own browser. Nothing is sent anywhere; no client data leaves the
device.

## Why it needs your approval

Modern SharePoint Online disables custom scripts by default (the "NoScript"
setting), which is correct and we are **not** asking you to change it. The
supported way to run custom code on a modern SharePoint page is the
**SharePoint Framework (SPFx)**, which runs in the normal page sandbox without
weakening tenant security. SPFx solutions must be **approved once in the
SharePoint App Catalog** by an administrator — that approval is what I'm
requesting.

## What you'd be approving

A single SPFx package file: **`div296-calculator.sppkg`** (built from the source
in `docs/hosting/spfx/` in our repo — happy to walk you through it or have it
reviewed first).

Security posture for your review:
- **No `requestAccessTokens` / no API permissions** requested in the package
  manifest — it talks to nothing.
- **`skipFeatureDeployment: true`** so it can be tenant-scoped and you control
  where it's added.
- No third-party CDNs — assets are bundled and served from our own app catalog.
- Source is in-house and auditable; not pulled from an external marketplace.

## The admin steps (one-time)

1. **Upload** `div296-calculator.sppkg` to the **App Catalog** site
   (`/sites/appcatalog` → *Apps for SharePoint*).
2. When prompted, choose whether to **make it available to all sites** (or just
   add it to specific sites).
3. **Trust** the solution when asked.
4. (Optional, for the Teams tab) approve it in the **Microsoft 365 admin
   center → Integrated apps** if you want it pinned in Teams.

After that, I can self-serve: add the web part to our team's SharePoint page and,
if wanted, add it as a **Website/app tab** in our Teams channel. It'll be
visible only to staff who already have access to that site/team — no public
exposure (the tool is proprietary).

## If you'd rather not approve a package

The fallback is fine too: I can host the same file on **Azure Static Web Apps**
(free tier, locked to our Entra ID tenant) and we embed that URL as a Teams
**Website tab**. That needs an Azure subscription rather than an app-catalog
approval — let me know which path you prefer.

---

*The tool is illustrative only and carries a "not financial/tax/legal advice"
disclaimer throughout. Questions: Aiden Hiew.*
