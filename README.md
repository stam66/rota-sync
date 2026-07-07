# rota-sync

Serves per-person iCalendar feeds (and JSON, for a future mobile PWA) from the
Echo and HF rota Google Sheets, so each person's rota commitments appear in
their nhs.net Outlook calendar.

- The Google Sheets stay the single source of truth — nothing changes about
  how the rotas are edited.
- Colleagues need **no** Google account and no login: each person gets a
  personal URL (protected by a random token) which they subscribe to once in
  Outlook on the web.
- `SK` appears in both rotas, so one feed merges Echo + HF commitments.
- Person-column statuses (`A/L`, `OOO`, `MBA`, …) also appear in that
  person's feed, with consecutive days merged into one multi-day event.
- Events are marked free (`TRANSP:TRANSPARENT`) so they don't block
  free/busy — flip that line in `buildIcs()` if you'd rather they show busy.

## Setup (one-off, ~10 minutes)

1. While signed in to the Google account that owns the rota sheets, go to
   <https://script.new> (creates a new Apps Script project).
2. Name it (e.g. `rota-sync`), delete the placeholder code, paste in
   `Code.gs`, save.
3. In the toolbar function dropdown pick **`setup`** and click **Run**.
   Grant the permission prompts (it only asks for Spreadsheets access).
   The log should report how many events it parsed.
4. Pick **`testParse`** and Run — check SK's next 20 events look right.
5. **Deploy** → *New deployment* → type **Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Copy the deployment URL (`https://script.google.com/macros/s/…/exec`).
6. Project settings (gear icon) → **Script properties** → add property
   `WEB_APP_URL` = the deployment URL.
7. Pick **`listFeedUrls`** and Run — the log prints every person's personal
   feed URL (plus an `ALL` feed). Paste one into a browser: you should get an
   `.ics` file.

## Each colleague then does this once

In Outlook on the web (nhs.net): **Calendar → Add calendar → Subscribe from
web**, paste their personal URL, give it a name (e.g. "Echo rota"). Done —
it updates automatically from then on.

> Test this step with your own nhs.net account first: NHSmail tenancy policy
> *usually* allows internet calendar subscriptions in OWA, but confirm before
> rolling out. Note Outlook refreshes subscribed calendars on its own
> schedule (typically every few hours) — fine for a rota, but last-minute
> swaps may lag. Emailed invite updates can cover that later if needed.

## Performance

Parsed events are cached for 15 minutes; run **`setupCacheWarmer()`** once
to install a 10-minute trigger that keeps the cache warm, so user requests
never pay for a full spreadsheet re-parse (which can take 30s+ on a bad
day). The cache is shared with the deployed web app — no redeploy needed.
Sheet edits reach the feeds within ~10 minutes.

## Weekly MediRota change report

Every Monday morning the script diffs the rotas against a snapshot taken at
the last report and emails the admin what changed — duty swaps (`SK → KG`),
additions, removals, leave changes (compressed into date ranges) — plus any
current rostered-while-away conflicts. The admin updates MediRota from the
email instead of hunting through the sheets.

Setup:

1. In `CONFIG.report.recipients`, add the admin's address(es).
2. Run **`setupChangeReport()`** once — grants Mail/Drive permission, saves
   the baseline snapshot (`rota-sync-snapshot.json` in your Drive) and
   installs the Monday 7am trigger. No web-app redeploy needed: triggers
   run the latest saved code.
3. Optional: **`previewChangeReport()`** logs the pending diff without
   emailing or advancing the snapshot; **`sendChangeReport()`** runs the
   real thing on demand.

Notes: nothing is emailed when there are no changes and no conflicts
(`sendIfEmpty: false`); the snapshot only advances after a successful send,
so a failed week's changes roll into the next report. `lookBackDays` (14)
controls how far back retroactive edits are caught.

## Invite mode — events in the person's REAL work calendar

The subscribed ICS feed lives in a separate overlay calendar in Outlook.
If someone wants rota shifts inside their primary nhs.net calendar instead,
opt them into **invite mode**:

1. Add their initials → nhs.net address in `CONFIG.invites.emails`.
2. Set the script project timezone to Europe/London (Project settings).
3. Run `setupInviteCalendar()` once — creates a "Rota sync" Google Calendar
   and installs a daily 6am `syncInvites()` trigger.
4. Run `syncInvites()` to send the first batch.

How it behaves:

- Each upcoming shift (next `windowDays` days) becomes an event on the sync
  calendar with the person as guest; Google emails a normal invitation,
  which Outlook auto-adds to their primary calendar (tentative until they
  accept).
- When the sheet changes, the daily sync deletes the stale event (guest
  gets a cancellation) and sends a fresh invite.
- Only duty shifts are sent — leave/status entries are not.
- The first sync sends one email per shift in the window, so start with a
  small `windowDays` (default 28) and raise it once people are happy.
  Google also rate-limits external guest invites, so avoid opting in many
  people and a huge window on the same day.

Feed and invites coexist fine — someone can have both, or either.

## Things to tune in `CONFIG` (top of Code.gs)

- **Role labels and times** — roles default to all-day events. `ECHO cover`
  is 08:00–17:00; `ECHO on call` is 17:00–08:00 overnight on weekdays and a
  24h shift (08:00–08:00) on weekends and bank holidays (`weekend24h: true`;
  BH detected from the Events column). Adjust `start`/`end` per role.
- **People** — the initials lists must match the per-person column headers.
  Add people there when the sheets gain columns.
- **New year tabs** are picked up automatically (any tab named like `2027`).
- **Feed window** — `daysBack` / `daysAhead`.

## Updating the code later

Edit the file in the Apps Script editor, then **Deploy → Manage deployments →
edit (pencil) → Version: New version → Deploy**. Reusing the same deployment
keeps everyone's URLs unchanged. (Creating a *new* deployment instead would
change the URL and break subscriptions.)

## Security notes

- Feed URLs are unauthenticated but unguessable (per-person token derived
  from a secret in Script Properties). Treat a URL like a password; if one
  leaks, change the `SECRET` script property — all tokens rotate, reissue
  URLs via `listFeedUrls`.
- Feeds contain only dates, role names, and initials — no patient data.

## Mobile PWA (`pwa/`)

A phone-friendly app over the same feed: **Me** (your upcoming commitments,
vertical list, today first), **Day** (who's on for any date — roles listed
vertically per rota, leave shown, duty-while-on-leave clashes flagged red),
**People** (anyone's rota), offline cache, installable to the home screen.

Hosting (free, once):

1. Create a GitHub repo, push the `pwa/` folder contents.
2. Repo Settings → Pages → deploy from branch → root. Note the
   `https://<user>.github.io/<repo>/` URL. (Cloudflare Pages works equally.)

Per-person setup is one link. Either send the page URL and let them paste
the **ALL feed URL** on the setup screen, or bake it in so one tap
configures everything:

    https://<user>.github.io/<repo>/#url=<URL-encoded ALL feed URL>&me=SK

The app stores the feed URL and identity locally, fetches a rolling
−14/+120-day window (`from`/`to` params on the JSON endpoint), and caches
the last data for offline use. iPhone: Share → Add to Home Screen.
Android: browser menu → Install app.

Local preview without a feed: open the page and tap **Demo**.

## Roadmap

1. **ICS feeds into Outlook** ✅ deployed; nhs.net OWA subscription confirmed
   working (overlay view merges it with the work calendar).
2. **Mobile PWA** ✅ built (`pwa/`) — awaiting hosting + rollout.
3. **MediRota change report** — weekly diff email to the admin ("these cells
   changed since last sync") so updating MediRota takes minutes.
