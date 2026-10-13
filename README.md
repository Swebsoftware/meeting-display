# Office availability display

A full-screen React/TypeScript PWA for a tablet outside an office. Cloudflare Workers hosts the static Vite/Tailwind UI and a same-origin Microsoft Graph API. No database is required. Demo mode is enabled by default; **turn it off before production**.

## Local development

Use Node 22 or newer and npm. From the repository root:

```sh
npm ci
cp .env.example .dev.vars
npm run build
npm run dev:api
# In a second terminal
npm run dev
```

Open the Vite address shown in its output. The frontend proxies `/api` to Wrangler on port 8787. The demo controls let you review Available, Busy, Meeting, Away, Do Not Disturb, and Offline. `?state=in_meeting` selects an initial demo state. Demo state parameters are ignored in live mode. The initial page waits for the API rather than assuming availability.

`npm run build` checks frontend and Worker TypeScript and builds production assets. `npm test` runs normalization, API, freshness, and daylight-saving tests. `npx playwright install chromium` installs the test browser; `npx playwright test` runs tablet layout and polling/error browser tests against the two development servers (start both first). `npm run preview` previews built frontend assets only; keep the Worker running and use Wrangler on port 8787 for a full production smoke test.

## Microsoft Entra setup — manual steps

1. In Microsoft Entra admin center, select your Microsoft 365 tenant and **App registrations → New registration**. Register a single-tenant application named Office Display. Client credentials is a server-to-server flow; **no redirect URL or interactive sign-in is required**.
2. Record the Directory (tenant) ID and Application (client) ID.
3. Under **API permissions → Add a permission → Microsoft Graph → Application permissions**, add `Presence.Read.All`, `Calendars.ReadBasic.All`, and `MailboxSettings.Read`. Grant tenant administrator consent for all three. Presence requires a Teams-enabled target user; calendar/mailbox calls require an Exchange Online mailbox. Delegated permissions are not used.
4. Under **Certificates & secrets**, create a client secret. Store the secret **value**, not its ID, securely. Set a rotation reminder before expiry. For higher-security deployments, replace the credential provider with a certificate-based assertion; v1 uses a server-side client secret.
5. Set `TARGET_USER_ID` to the target user's object ID or user principal name. No `User.Read.All` permission is needed because the app never searches the directory.
6. For local live testing, set `DEMO_MODE=false` and fill the four Microsoft variables in the ignored `.dev.vars` file. Never commit it or put secrets into `VITE_*` variables.
7. For Cloudflare, follow the deployment steps below. Verify `/api/status` and the tablet with a real meeting, away status, and DND.

Application permissions can grant broad tenant access. Ask your Exchange administrator to scope calendar/mailbox access to the target mailbox using supported Exchange application access controls / Application RBAC. Teams presence authorization is separate. Confirm effective permissions with your tenant administrator.

Official references: [Get presence](https://learn.microsoft.com/en-us/graph/api/presence-get), [Calendar view](https://learn.microsoft.com/en-us/graph/api/user-list-calendarview), [Mailbox settings](https://learn.microsoft.com/en-us/graph/api/user-get-mailboxsettings), [Client credentials](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow). Live Microsoft integration requires tenant credentials and was not validated using this repository's demo checks.

## Cloudflare deployment

```sh
npx wrangler login
npx wrangler secret put MICROSOFT_TENANT_ID
npx wrangler secret put MICROSOFT_CLIENT_ID
npx wrangler secret put MICROSOFT_CLIENT_SECRET
npx wrangler secret put TARGET_USER_ID
npm run deploy
```

Before the final command, change `DEMO_MODE` to `"false"` in `wrangler.jsonc`. Remove `DISPLAY_TIME_ZONE` from the Worker vars to use the user's Outlook mailbox timezone, or explicitly set an IANA override such as `America/New_York`. The example override is for demo use. Windows mailbox timezone names are converted to IANA via `windows-iana`. Calendar responses request UTC from Graph and are converted only when displayed. Date, clock, and meeting end time use the same zone.

Workers Assets serves `dist`; the Worker handles `/api/*`. The response includes status, displayStatus, message, busyUntil (nullable ISO UTC), updatedAt, timeZone, and demo. Graph tokens are kept only in Worker memory until expiry, with a one-minute renewal margin. Status is cached per isolate for 20 seconds and concurrent requests share a pending request. No access tokens, event titles, attendees, descriptions, or raw Graph errors reach the frontend. Network destinations: `login.microsoftonline.com` and `graph.microsoft.com`.

Place the deployment behind **Cloudflare Access** with a policy permitting the tablet / approved users if availability information is private. Authenticate the tablet using the supported Access browser session and configure session duration appropriately. Do not embed an Access service token in the PWA. Protecting only the HTML leaves the API public: protect the entire hostname. No login UI is built into this single-office application. Restrict CORS to same origin (no cross-origin headers are emitted). Secret rotation uses the same `wrangler secret put` command and redeploy as needed.

## Status rules

Highest priority wins:

1. Presenting or DoNotDisturb → Do Not Disturb.
2. InACall or InAMeeting → In a Meeting.
3. A current non-cancelled, non-all-day calendar entry marked busy or out-of-office → In a Meeting.
4. Busy → Busy.
5. Away or BeRightBack → Away.
6. Available → Available.
7. Everything else → Status unavailable.

Free/tentative/working-elsewhere and all-day events do not reserve the door. Calendar busy entries override even offline presence by design; an event is an availability reservation, not proof of physical attendance. Calls are labeled In a Meeting for the simple door interface. When a busy calendar entry overlaps the current moment, show its end time; overlapping entries use the latest end. A Teams meeting without a matching calendar reservation has no invented end time. This is an estimate and can differ from actual call duration.

Calendar queries cover yesterday through tomorrow and request only `start,end,showAs,isCancelled,isAllDay`. A paginated result is rejected conservatively rather than silently showing incomplete availability. Exceptionally busy calendars may need pagination support in a later version. Full presence/calendar/mailbox retrieval must succeed before a new status is issued; failures return 503 with Retry-After.

The UI polls every 25 seconds with a 12-second request timeout. It keeps the last successful observation for at most two minutes, then shows Status unavailable. A separate freshness timer expires stale information even while a request is pending. Returning to the tab triggers an immediate refresh. No stale API data is stored by the service worker. This uses polling rather than Graph webhook subscriptions and needs no durable subscription database. Multiple Workers isolates/tablets may increase Graph usage; monitor throttling and add shared caching if scaling beyond a few tablets.

## Tablet / PWA operation

Use HTTPS in production, install from the browser's install / Add to Home Screen action, then configure your tablet's kiosk or guided-access mode. The app requests Screen Wake Lock and reacquires it on visibility changes where supported. Browsers may release the lock; configure OS sleep policy as well. The Full screen control requires a user gesture and browser support. Landscape and portrait layouts work without scrolling at standard tablet sizes. Normal operation has no menus; demo controls disappear in live mode.

The service worker caches previously loaded UI assets for basic offline use but never caches `/api`. Offline startup shows unavailable until fresh data arrives. First launch requires connectivity. SVG and raster icons (192/512 and Apple touch icon) are supplied. PWA install support varies by platform.

## Structure

- `src/main.tsx` — tablet UI, polling, freshness, wake lock, fullscreen.
- `src/status.ts` — normalized contract, labels, formatting and validation.
- `src/style.css` — responsive sign design, Tailwind entry point.
- `worker/index.ts` — API routing, demo responses, cache and safe errors.
- `worker/graph.ts` — private authentication, Graph queries, timezone resolution.
- `worker/normalize.ts` — presence/calendar priority logic.
- `public/` — PWA manifest, icon and offline service worker.
- `tests/` — unit/API and browser workflow tests.

The normalized API boundary allows future logos, photos, multiple users and configurable messaging without exposing Graph data to the UI. Meeting titles are intentionally not fetched; an explicit future server-side privacy configuration would be needed before requesting/exposing them.
