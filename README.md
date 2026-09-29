# teams-transcript-capture

A Chrome MV3 extension (load-unpacked) that captures Microsoft Teams meeting transcripts as Markdown.

When you open a meeting's **Transcript** tab, the Teams client renders the transcript inside a cross-origin SharePoint Stream iframe. This extension runs a MAIN-world content script in that iframe, intercepts the SharePoint network response that carries the transcript metadata, fetches the full transcript JSON (or VTT fallback), converts it to a clean speaker-grouped Markdown document, and gives you one-click **Copy** or **Download**.

## Layout

- `manifest.json` — MV3 manifest. Content scripts are dynamically registered only while capture is enabled. Hosts: `*.teams.microsoft.com` (future use, no content scripts), `*.sharepoint.com`.
- `intercept.js` — MAIN-world content script injected at `document_start` on `*.sharepoint.com`. Monkey-patches `fetch` / `XMLHttpRequest` to observe transcript metadata and pass it to the isolated world.
- `content.js` — isolated-world content script at `document_idle` on `*.sharepoint.com`. Receives metadata, fetches the transcript (`?format=json`), converts to Markdown, and drives the floating UI.
- `src/normalize.js` — normalizes the three metadata JSON shapes into `{temporaryDownloadUrl, displayName, languageTag}`.
- `src/time.js` — timestamp parsing/formatting helpers.
- `src/markdown.js` — JSON → Markdown and VTT → Markdown converters.
- `src/ui.js` — shadow-DOM floating panel.
- `background.js` — service worker. Handles `chrome.downloads` (data URI for small payloads, offscreen Blob for large ones) and a CSP-fallback fetch bridge.
- `offscreen.html` / `offscreen.js` — offscreen document used for `Blob`/`URL.createObjectURL`, which the service worker cannot access.

## Tests

```bash
npm install          # installs Playwright dev dependency
npx playwright install chromium
npm test             # 56 unit + 1 Playwright integration test
npm run test:unit
npm run test:integration
npm run lint:syntax
```

The integration test spins up a local HTTPS fixture server impersonating a SharePoint Stream page, loads the unpacked extension in Chromium, and asserts the full intercept → fetch → convert → download pipeline end-to-end.

## Sovereign-cloud limitation

This extension targets `https://*.teams.microsoft.com/*` and `https://*.sharepoint.com/*` only. Sovereign-cloud tenants (e.g. GCC / GCC-High / DoD on `*.teams.microsoft.us`, `*.sharepoint.us`, and similar) are **not** covered by the current `host_permissions` / `matches` and will not be intercepted until explicit sovereign host patterns are added.

## Loading (dev)

1. Open `chrome://extensions`, enable **Developer mode**.
2. Click **Load unpacked** and select this directory.
3. Right-click the extension's toolbar icon and choose **Enable Transcript Capture**. Reload the SharePoint transcript page after enabling.
4. Open a Teams meeting and click the **Transcript** tab. A floating panel appears in the bottom-right of the transcript pane, transitions to **"Transcript ready: N entries"**, then click **Copy .md** or **Download .md**.

## Enable / disable capture

Capture is disabled by default, so no extension code loads into SharePoint pages until you explicitly enable it. Right-click the extension's toolbar icon and choose **Enable Transcript Capture**; reload the transcript page to start capture. Use the same menu and choose **Disable Transcript Capture** when you are finished. The setting persists across browser restarts. A script that is already running in an open page remains there until that page is reloaded, so reload after either change for it to take effect immediately.
