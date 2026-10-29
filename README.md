# Office availability display

A full-screen React/TypeScript PWA for a tablet outside an office. Cloudflare Workers hosts the static Vite/Tailwind UI and a same-origin Microsoft Graph API. No database is required. The deployed app uses live Microsoft status. Use `DEMO_MODE=true` in `.dev.vars` for local demo development.

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

## Connect Microsoft — recommended setup

The `/connect` page signs the configured office user into Microsoft using a tenant-specific authorization-code flow with PKCE. No tokens reach frontend JavaScript. Each person receives a distinct display link; one user signing in never replaces another user’s connection. The door sign keeps its minimal layout; connection settings live on a separate page.

### Microsoft Entra manual steps

1. Register a **single-tenant** app, or reuse your Office Display registration.
2. Under **Authentication**, add a **Web** redirect URI exactly matching `https://meeting-display.falling-sea-3e49.workers.dev/api/auth/callback`. Do not select SPA, implicit grant, or public-client flows. For a different deployment hostname, update both this URI and `APP_ORIGIN`.
3. Under **API permissions → Microsoft Graph → Delegated permissions**, add `User.Read`, `Presence.Read`, `Calendars.ReadBasic`, `MailboxSettings.Read`, and `offline_access`. To avoid individual admin-approval requests, an authorized administrator can grant tenant-wide consent once for these delegated permissions. The organization’s user-consent, assignment, Conditional Access, and MFA policies still apply; the application cannot bypass them. Application permissions are not required for this flow; remove unused application permissions after confirming no other system depends on them. Your organization's consent policies may require **Grant admin consent**. The target user needs Teams and an Exchange Online mailbox.
4. Under **Certificates & secrets**, create a client secret. Securely save the **Value**, not the secret ID, in Cloudflare → Workers & Pages → meeting-display → Settings → Variables and Secrets as `MICROSOFT_CLIENT_SECRET`. Never paste it into chat, commit it, or use a `VITE_*` variable. Set a secret rotation reminder.
5. Open `/connect` on the deployed hostname and click **Connect Microsoft**. Sign in with your own Microsoft 365 work account in this organization. Each account gets a private settings session and a unique `/display/<random-id>` link. Copy that link to its tablet. Reconnecting the same account retains its link. Accounts from other organizations cannot use this single-tenant Entra registration. `TARGET_USER_ID` is retained only for legacy storage compatibility; it does not select a default display or restrict sign-ups.
6. Confirm connection success, then set `DEMO_MODE=false` and remove `DISPLAY_TIME_ZONE` to use Outlook's timezone. Redeploy. Test Available, Busy, DND, a real calendar meeting, and token renewal. Tenant-specific live validation requires the user's sign-in and consent.

### Backend configuration

| Setting | Purpose |
| --- | --- |
| `AUTH_MODE=delegated` | Microsoft user sign-in (recommended) |
| `MICROSOFT_TENANT_ID` | Entra directory ID, non-secret |
| `MICROSOFT_CLIENT_ID` | Entra application ID, non-secret |
| `MICROSOFT_CLIENT_SECRET` | Worker secret for confidential-client token exchange |
| `TARGET_USER_ID` | Original account ID / email for the legacy root display only |
| `APP_ORIGIN` | Fixed HTTPS app origin; callback is `/api/auth/callback` |
| `TOKEN_ENCRYPTION_KEY` | Worker secret: base64url-encoded random 32-byte AES key |
| `CONNECTIONS` | Cloudflare KV binding for encrypted tokens |
| `DEMO_MODE` | `true` keeps demo visible; `false` enables live Graph status |
| `DISPLAY_TIME_ZONE` | Optional IANA override; omit to respect mailbox timezone |

Cloudflare KV retains encrypted connections per user and a tenant/user ID index for stable links across Worker restarts. The settings page uses an encrypted 30-day HttpOnly session cookie bound to one connection; no unauthenticated user can retrieve another person’s link through the settings API. Signing out clears only that settings session; the tablet continues to work. Access and refresh tokens are encrypted with AES-256-GCM and a random IV before storage. OAuth flow data uses an encrypted 10-minute HttpOnly, Secure, SameSite=Lax `__Host-` cookie; state, PKCE, same-origin POST initiation, a fixed callback, tenant-specific endpoints, and Graph-confirmed account identity protect the connection. The callback fetches only account identity and stores no meeting content. Auth routes are no-store and omit raw Microsoft errors.

The backend refreshes tokens automatically with a one-minute expiry margin, retaining rotated refresh tokens. Revoked authorization / interaction-required responses prompt reconnection. Reconnect at `/connect` after a password/consent policy change. KV may take up to 60 seconds to propagate a newly connected account. Refreshes and status caches are keyed separately for every connection. Refreshes are deduplicated per user per isolate; KV is not a distributed lock. Microsoft refresh tokens remain usable across concurrent refreshes. For a high-volume deployment, use a per-account Durable Object to serialize refreshes. Concurrent first-time sign-ins for the same account may observe delayed KV index propagation; retry sign-in from the original browser if the settings page has not appeared yet.

**Do not rotate the encryption key without migrating tokens**: existing ciphertext cannot be decrypted with a new key, so the account must reconnect. The generated key is stored only as a Cloudflare secret. No browser session is required on the tablet after the account is connected. Each personal link contains a cryptographically random 256-bit capability. Anyone with that link can view normalized availability without signing in. Treat it as private and do not post it publicly. Display responses use no-referrer and noindex headers; the service worker does not store personal URLs or API responses. The main URL opens connection settings and never displays a default account. Live status requires a personal display link. Cloudflare Access can additionally protect the hostname. Changing encryption keys invalidates settings sessions as well as tokens.

### Cloudflare deployment

```sh
npx wrangler login
# For a fresh deployment: create KV and insert its ID into wrangler.jsonc.
npx wrangler kv namespace create CONNECTIONS
npx wrangler secret put MICROSOFT_CLIENT_SECRET
# Generate a random 32-byte base64url key securely; save it as:
npx wrangler secret put TOKEN_ENCRYPTION_KEY
npm run deploy
```

The existing deployment already has the KV binding and encryption secret; do not recreate them. Configure public IDs and app origin under `vars` in `wrangler.jsonc`. Local development uses `.dev.vars`, Wrangler's local KV, and a separate test app / HTTPS tunnel with a registered callback for live OAuth; local demo works without credentials. `APP_ORIGIN` must be HTTPS because OAuth cookies are Secure. Do not use production KV with local development.

Place the entire hostname behind **Cloudflare Access** if status is private. Protect both HTML and API and authenticate the tablet using an Access browser session; do not embed service tokens in the PWA. All Microsoft authentication/API requests originate on the Worker (`login.microsoftonline.com`, `graph.microsoft.com`). Cloudflare runtime egress is independent of development workspace networking.

### Alternative: application credentials

`AUTH_MODE=application` retains the original unattended client-credentials flow. It requires **application** permissions `Presence.Read.All`, `Calendars.ReadBasic.All`, and `MailboxSettings.Read`, all with administrator consent. Configure the same tenant/client/secret/target fields as Worker secrets or vars. No callback, KV, encryption key, or Microsoft user sign-in is used. Ask your Exchange administrator to scope mailbox access using supported application access controls / Application RBAC. Teams authorization is separate. Do not combine both modes' permission grants unnecessarily.

### Timezones and privacy

Windows Outlook timezone names are converted to IANA via `windows-iana`. Calendar responses request UTC; dates, clock, and busy-until values render in the mailbox timezone unless explicitly overridden. The API returns normalized status, displayStatus, message, busyUntil, updatedAt, timeZone, and demo. It never returns tokens, meeting titles, descriptions, attendees, or raw Graph errors.

Official references: [Presence](https://learn.microsoft.com/en-us/graph/api/presence-get), [Calendar view](https://learn.microsoft.com/en-us/graph/api/user-list-calendarview), [Mailbox settings](https://learn.microsoft.com/en-us/graph/api/user-get-mailboxsettings), [Authorization-code flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow).

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

The UI polls every 10 seconds with a 12-second request timeout and skips overlapping requests. The backend caches each user’s status for 5 seconds; typical update latency is 10–15 seconds after Graph reflects a change, plus request time. Microsoft propagation delays can add more time. It keeps the last successful observation for at most two minutes, then shows Status unavailable. A separate freshness timer expires stale information even while a request is pending. Returning to the tab triggers an immediate refresh. No stale API data is stored by the service worker. This uses polling rather than Graph webhook subscriptions and needs no durable subscription database. Multiple Workers isolates/tablets may increase Graph usage; monitor throttling and add shared caching if scaling beyond a few tablets.

## Tablet / PWA operation

Use HTTPS in production, install from the browser's install / Add to Home Screen action, then configure your tablet's kiosk or guided-access mode. The app requests Screen Wake Lock and reacquires it on visibility changes where supported. Browsers may release the lock; configure OS sleep policy as well. The Full screen control requires a user gesture and browser support. Landscape and portrait layouts work without scrolling at standard tablet sizes. Normal operation has no menus; demo controls disappear in live mode.

The service worker caches previously loaded UI assets for basic offline use but never caches `/api`. Offline startup shows unavailable until fresh data arrives. First launch requires connectivity. SVG and raster icons (192/512 and Apple touch icon) are supplied. PWA install support varies by platform.

## Structure

- `src/main.tsx` — tablet UI, polling, freshness, wake lock, fullscreen.
- `src/status.ts` — normalized contract, labels, formatting and validation.
- `src/style.css` — responsive sign design, Tailwind entry point.
- `worker/index.ts` — API routing, demo responses, per-user cache and safe errors.
- `worker/auth.ts` — OAuth, target-account verification, encrypted storage, token refresh.
- `src/Connect.tsx` — separate Microsoft connection page.
- `worker/graph.ts` — private authentication, Graph queries, timezone resolution.
- `worker/normalize.ts` — presence/calendar priority logic.
- `public/` — PWA manifest, icon and offline service worker.
- `tests/` — unit/API and browser workflow tests.

The normalized API boundary allows future logos, photos and configurable messaging without exposing Graph data to the UI. Meeting titles are intentionally not fetched; an explicit future server-side privacy configuration would be needed before requesting/exposing them.

## Sharing with coworkers

Send coworkers `https://meeting-display.falling-sea-3e49.workers.dev/connect`. They sign in using their own Microsoft 365 account, then select **Copy link** or **Open my display**. No new app registration, tenant ID entry, client secret entry, or deployment is required per user. A tenant administrator should grant organization-wide consent once if policies otherwise require individual approval requests. The app remains single tenant; supporting other organizations would require an explicitly reviewed multi-tenant registration and tenant validation.

The root URL opens Microsoft connection settings. Open your personal link from `/connect` to use a tablet display. To retrieve a personal link for the original account, sign in again on `/connect`; your existing personal link stays the same. Recovery uses Microsoft sign-in on `/connect`; do not share an authenticated settings browser session.

Personal displays load a per-link PWA manifest so an installed tablet app opens the correct user’s URL rather than the legacy root. Personal URLs and their manifests are not stored by the service worker. For a separate browser test API, run `npx wrangler dev --port 8788 --var DEMO_MODE:true`, then `TEST_API_ORIGIN=http://127.0.0.1:8788 CHROMIUM_PATH=/usr/bin/chromium npx playwright test` while Vite runs.

The top-left label uses the connected user’s Microsoft `jobTitle` (via the existing delegated `User.Read` permission), cached per user for five minutes. If no title is set or the optional profile lookup fails, it displays OFFICE. Job title is included in the normalized public display response; meeting content remains private.
