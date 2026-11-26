# Sallee Load Board — quick re-entry notes

Repo: `canumanu/SALLEE-LOAD-BOARD` — live at `canumanu.github.io/SALLEE-LOAD-BOARD/`
React app, main component `SalleeLoadBoard.jsx`, entry `src/main.jsx`.

## How it's built/deployed
- Every `.jsx`/`src` edit requires rebuilding `bundle.js` before changes go live — the site serves `bundle.js` directly. Since 2026-10-15 the repo has a `package.json`: `npm install` then `npm run build` (esbuild, entry `src/main.jsx`).
- `index.html` (see `claude/index.html` in this project) is the static shell — dark-green `#root` div, then `config.js`, then `bundle.js`. All real UI comes from the bundle.
- `config.js` holds the sign-in/SharePoint settings (clientId, tenantId, site, list name, refresh). Edit without rebuilding. Empty `clientId` = demo mode (no sign-in, no New Order).
- Deployed on GitHub Pages. Azure AD app registration `sallee-load-board-sync` handles the sheet sync. External cron trigger (cron-job.org) fires the sync workflow — GitHub's native scheduler proved unreliable.
- Past gotchas already solved: `bundle.js` path must be relative (not absolute), a SharePoint tab name had a trailing space, and there were a couple of PAT security incidents.

## Data flow
1. SharePoint `ORDERS_LIST_2026.xlsx` → `sync_orders.py` (Microsoft Graph) → `data/orders.json` via GitHub Actions.
2. **Load Orders** SharePoint list (on the Sallee Dispatch Board site, `/sites/SALLEEDISPATCHBOARD`) — orders entered through the app's "+ New Order". The board reads it live after sign-in (every 30s); `sync_orders.py` also appends open list orders to `orders.json` as `sp-<itemId>` (with `source: "app"`) so the dispatch board's Browse sees them. Live list copy wins over the orders.json copy; completed/canceled list orders are hidden immediately.

Real sheet column headers: `ORIGIN FARM`, `ORIGIN TRAINER`, `ORIGIN TRACK`, `DESTINATION TRACK`, `DESTINATION TRAINER`, `DESTINATION FARM`, `STALL SPACE`.

- `ORDERS` sheet supports comma-separated multi-horse orders via `HORSE NAME(S)` and `STALL SPACE` columns (e.g. `"QUIRKY 25, MISS VEKOMA 25"` / `"1.5, 3"`), parsed into a per-order `horses` array. Stall values work even if horse names are blank.
- `CANCELED` column (renamed from `CANCEL`) excludes orders the same way `COMPLETED` does.
- Load Orders list columns (all text): `RequestedBy, RequestedDates, DateTaken, TakenBy, Origin, OriginTrack, OriginFarm, OriginTrainer, Destination, DestinationTrack, DestinationTrainer, DestinationFarm, HorseNames, StallSpace, Notes, TripDate, Drivers, Status` (Status = Open / Completed / Canceled). Matched by display name at runtime, so internal names (e.g. `field_3` from an Excel import) don't matter.
- Board trims stray spaces on string fields so "KENTUCKY " and "KENTUCKY" group together.

## Key features
- **Intro + Microsoft 365 sign-in (built 2026-10-15):** same flow as the shop board — intro animation (van's ramp drops, horse in a green/gold sheet walks up and in, ramp shuts, van drives off) → "Sign in with Microsoft" → "Good afternoon, <name>" → board. Intro plays once per tab session, skippable. Reuses the dispatch board's "Sallee Dispatch Board Sign-in" app registration (`3530a45a-…`, delegated `Sites.ReadWrite.All`); MSAL cache in localStorage so people stay signed in. Sign out = this board only.
- **+ New Order (built 2026-10-15):** form with requested by, taken by (defaults to signed-in first name), requested date(s) (free text + date picker + NEXT AVAILABLE/ASAP/READY chips), From/To region, track, trainer, farm, horses (name + 1.5 stall / 1 box / other), notes. Saves to the Load Orders list. App orders show "✎ Edit / close out" in the load modal (edit, trip date, drivers, Mark completed, Cancel). Spreadsheet orders are labeled "from spreadsheet" and stay read-only.
- Permissions come from SharePoint: Load Orders list has its own permissions (planned: Everyone except external users → Contribute). 403 on save shows a "no permission" message.
- Airport-departures-board aesthetic, Sallee green/gold branding, origin-grouped pagination, scrolling ticker, stall map modal.
- `StallStrip` renders one colored block per horse, with a `┊` separator between horses sharing an order (labeled "1 box" / "1.5 stall").
- `LoadModal` has a "Get Directions" button using a `TRACK_ADDRESSES` lookup (~25 racetracks/training centers) plus a farm+region fallback, opening Google Maps.
- "Build a Load" tool, still backed by `localStorage`.
- NAV (Microsoft) is still the system of record for billing/BOL — this board is purely an operational/coordination layer.

## Go-live status (as of 2026-10-15)
Code done and tested against a mocked list; see `LOAD_ORDERS_SETUP.md` in the repo. Waiting on Manu to: upload files (Claude session couldn't push), create the Load Orders list from `setup/Load_Orders_list_template.xlsx`, set list permissions, add the two SPA redirect URIs (`/SALLEE-LOAD-BOARD/` and `/SALLEE-LOAD-BOARD/index.html`) to the dispatch sign-in app.

## Known gaps / later
- `data/orders.json` is a public file in a public repo — the sign-in gates the board, not that file. Fully locking it down would mean reading the sheet live via Graph and changing the dispatch board's Browse.
- **Print view:** notes from 2026-09-10 describe a standalone `print.html` + `buildPrintHref()`, but neither was in the repo as of 2026-10-15 — never uploaded, or lost. Needs rebuilding/uploading.
- "Build a Load" still localStorage-only; moving it to a shared list is the next multi-user step.
- Stray file `data/orders. Json` (with a space) in the repo can be deleted.

## Not this repo
Separate Sallee projects: **SALLEE-DISPATCH-BOARD** (column dispatch grid; MSAL + Drivers/Jobs lists on `/sites/SALLEEDISPATCHBOARD`; has its own van drive-in intro), **Sallee-shop-board** (shop TV; intro → sign-in → reads shop.xlsx via Graph), plus barn-boards, fuel board, event calendar, local dispatch board.
# Keeping SalleeLoadBoard.jsx and bundle.js in sync

**Found Nov 26, 2026:** the `SalleeLoadBoard.jsx` committed to the
`canumanu/SALLEE-LOAD-BOARD` repo had drifted badly out of sync with the
live `bundle.js` actually running on GitHub Pages. The committed source was
an older version from before sign-in and the "+ New Order" SharePoint
integration existed, while `bundle.js` had clearly been rebuilt (at some
point, in some other session) from a newer source that wired all of that
in — but that newer `SalleeLoadBoard.jsx` was apparently never uploaded
back to the repo, only the compiled `bundle.js`.

Net effect: the repo's plain-text source didn't reflect what was actually
live, which is risky for any future change (editing the stale source and
rebuilding would have silently deleted the sign-in header, "+ New Order",
mark-completed/cancel, etc. from the live site).

**What was done:** reconstructed the missing pieces by decompiling the
live `bundle.js` (useAppOrders polling hook, the spreadsheet/app-order
merge, the OrderForm wiring, per-order "Edit / close out" button inside
the Load Detail modal, the account/sign-out header bit, and the
mark-completed/cancel handlers) and merged them back into
`SalleeLoadBoard.jsx`, then rebuilt `bundle.js` from that reconciled
source. Also improved the mark-completed/cancel error handling —
previously any failure just said "(connection or permissions issue)";
now it surfaces the actual cause (expired sign-in, a 403/permissions
problem on the Load Orders list, a 404 if the item's gone, etc.) via a
new `describeGraphError()` helper in `SalleeLoadBoard.jsx`.

**Reported bug this was chasing:** marking an order "Completed" from the
order form silently did nothing on the first try (order stayed open, no
SharePoint change), and on a second try (after signing out/back in) gave
a generic "connection issue" alert. Most likely explanation: the first
click's Graph token had silently expired, triggering MSAL's
`acquireTokenRedirect` (a real page navigation that can look instant/
silent with an existing SSO session, but loses whatever UI action was
mid-flight); the second click was a genuine request that failed for a
real reason the old generic alert didn't reveal. The improved error
message should reveal the real cause next time this happens.

**Going forward:** whenever `SalleeLoadBoard.jsx` is edited and
`bundle.js` rebuilt, upload BOTH files together — never upload just a
rebuilt `bundle.js` on its own, or the repo's source will drift out of
sync with the live site again exactly like this.
