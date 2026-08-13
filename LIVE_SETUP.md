# Launching the Load Board as a live web page

This gets you from "zip file on your computer" to a real URL your
dispatchers can bookmark, kept in sync with the ORDERS sheet in
SharePoint automatically. No build step, no local install required —
everything runs on GitHub.

## What's in the folder

```
sallee-load-board-live/
├── index.html              ← the page itself
├── bundle.js                ← the board, already built and ready to serve
├── data/orders.json         ← the current snapshot (gets overwritten by the sync)
├── scripts/sync_orders.py   ← pulls fresh data from SharePoint
├── .github/workflows/
│   └── sync-orders.yml      ← runs the script on a schedule
└── LIVE_SETUP.md            ← this file
```

`index.html` and `bundle.js` are the compiled site — you don't need
Node, npm, or any build tool to publish them as-is. If you ever want
the board's look or behavior changed, send me the request and I'll
rebuild `bundle.js` for you.

## Step 1 — Create the GitHub repo and upload the files

1. On github.com, click **New repository**. Name it something like
   `sallee-load-board`. Public or private both work with GitHub Pages
   (private repos need GitHub Pro/Team/Enterprise for Pages — public is
   simplest if that's not already the case for your org).
2. Upload everything from the `sallee-load-board-live` folder into the
   repo root, keeping the folder structure exactly as it is (drag-and-drop
   on the repo's "Add file → Upload files" page works, or use `git push`
   if you're comfortable with git).

## Step 2 — Turn on GitHub Pages

1. In the repo, go to **Settings → Pages**.
2. Under **Build and deployment → Source**, choose **Deploy from a
   branch**.
3. Set **Branch** to `main` and the folder to `/ (root)`, then **Save**.
4. GitHub will show a URL at the top of that page once it's live —
   something like `https://your-org.github.io/sallee-load-board/`. It
   can take a minute or two the first time.

That's the whole publishing step. Visit that URL and you'll see the
board — running on the snapshot data in `data/orders.json` until the
live sync is connected.

## Step 3 — Let the Action commit back to the repo

The sync workflow needs permission to push its own commits (updating
`data/orders.json`).

1. Go to **Settings → Actions → General**.
2. Scroll to **Workflow permissions**.
3. Select **Read and write permissions**.
4. **Save**.

This is the single most common thing people miss — if it's left on
"Read only," the sync will run but fail silently on the commit step.

## Step 4 — Register an app in Azure AD

This is what lets the sync script read the file without a person
logged in.

1. Go to **portal.azure.com → Azure Active Directory → App
   registrations → New registration**. Name it something like
   `sallee-load-board-sync`. Leave the redirect URI blank.
2. Once created, note the **Application (client) ID** and **Directory
   (tenant) ID** from the app's Overview page.
3. Go to **Certificates & secrets → New client secret**. Copy the
   secret **value** immediately — it's only shown once.
4. Go to **API permissions → Add a permission → Microsoft Graph →
   Application permissions**, and add `Sites.Read.All`.
5. Click **Grant admin consent** for your organization (needs a
   SharePoint/Global admin — ask them if that's not you).

## Step 5 — Add the GitHub secrets

Back in your repo: **Settings → Secrets and variables → Actions → New
repository secret**. Add these four:

| Secret name            | Value                                              |
|-------------------------|-----------------------------------------------------|
| `AZURE_TENANT_ID`       | from Step 4                                         |
| `AZURE_CLIENT_ID`       | from Step 4                                         |
| `AZURE_CLIENT_SECRET`   | from Step 4                                         |
| `SP_SHARE_URL`          | the ORDERS_LIST_2026.xlsx SharePoint link, pasted exactly as copied |

## Step 6 — Run the sync once, manually

1. Go to the **Actions** tab → **Sync orders from SharePoint** →
   **Run workflow** → **Run workflow** (green button).
2. Wait ~30 seconds, then refresh. You should see a new commit updating
   `data/orders.json`.
3. Reload your published page (`https://your-org.github.io/...`). The
   badge in the top left should switch from **○ SNAPSHOT** to
   **● LIVE**, and the counts should match what's actually in the sheet.

If the run fails, click into it in the Actions tab — the error from
Microsoft Graph will be in the log. The two most common first-run
issues are a missing admin consent (Step 4.5) or the workflow
permissions not being set to read/write (Step 3).

## Step 7 — Let it run itself

Once the manual run works, the schedule in `sync-orders.yml` takes over
— every 5 minutes it checks SharePoint and commits any changes. Nobody
needs to touch anything for the board to stay current. New rows typed
into the ORDERS tab show up as a ticker alert and a highlighted row
within a few minutes of being entered.

## Adjusting things later

- **How often GitHub checks SharePoint**: edit the cron schedule in
  `.github/workflows/sync-orders.yml`.
- **How often the page checks for new data**: edit
  `POLL_INTERVAL_MS` near the top of `SalleeLoadBoard.jsx` — this only
  matters if you rebuild `bundle.js`, since the published site runs off
  the compiled file, not the source directly.
- **Trailer capacity, colors, layout, anything visual**: send me the
  change and I'll edit `SalleeLoadBoard.jsx` and hand you a freshly
  built `bundle.js` to drop back into the repo (just replace the file
  and commit — Pages picks it up automatically).
