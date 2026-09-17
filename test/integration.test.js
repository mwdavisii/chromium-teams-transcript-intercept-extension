// test/integration.test.js
// End-to-end QA harness for the teams-transcript-capture Chrome extension.
//
// What it does:
//   1. Starts the local HTTPS fixture server (test/server.js) on 127.0.0.1:9876
//      serving a `*.sharepoint.com`-hosted Stream page, a transcript-metadata
//      JSON, and a transcript-content JSON (all same-origin).
//   2. Launches Chromium (via Playwright) with the unpacked extension loaded
//      and `--host-resolver-rules` mapping `fakefixture.sharepoint.com` to
//      127.0.0.1 so the extension's `https://*.sharepoint.com/*` content scripts
//      actually fire on the local fixture page.
//   3. Navigates to the fixture page, whose own fetch() hits the metadata
//      endpoint; the MAIN-world hook intercepts it and the isolated-world
//      content script captures + converts the transcript.
//   4. Waits up to 30s for the shadow-DOM panel to show "Transcript ready".
//   5. Clicks "Download .md", waits for the `.md` file to land on disk, and
//      asserts >=10 `### ` speaker headings, >=10 `_[` timestamp markers, and
//      >=2 distinct speakers.
//
// Skipping (not failing): the test skips with a logged reason when Playwright
// is not installed, when no Chromium binary is available, or when the browser
// cannot be launched.
//
// Run:  node --test test/integration.test.js     (from the repo root)

'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');

const REPO_ROOT = path.resolve(__dirname, '..');
const { startServer, stopServer, PORT, HOSTNAME } = require('./server');

// ---------------------------------------------------------------------------
// Availability detection (so the test can skip instead of fail).
// ---------------------------------------------------------------------------

let chromium;
try {
  chromium = require('playwright').chromium;
} catch (e) {
  chromium = null;
}

const CANDIDATE_CHROMIUM_PATHS = [
  process.env.TTC_CHROMIUM_PATH,
  '/usr/bin/chromium',
  '/usr/bin/chromium-browser',
  '/opt/google/chrome/chrome',
  '/usr/bin/google-chrome-stable',
].filter(Boolean);

function findChromiumExecutable() {
  for (const p of CANDIDATE_CHROMIUM_PATHS) {
    try {
      if (fs.existsSync(p)) return p;
    } catch (e) {
      // ignore
    }
  }
  return null;
}

const CHROMIUM_PATH = findChromiumExecutable();
const PAGE_URL = `https://${HOSTNAME}:${PORT}/stream.html`;

// Playwright has no built-in way to capture a service-worker-initiated
// chrome.downloads.download (the download is unattributed to any page), so we
// watch the dedicated download directory directly instead of
// page.waitForEvent('download'). Note: for a data: URI download under CDP
// `Browser.setDownloadBehavior`, Chrome writes the file under a bare UUID
// name with NO `.md` extension (see learnings note), so we match ANY new file,
// then confirm it's markdown by its content.
function listFiles(dir) {
  const found = [];
  function walk(d) {
    let entries;
    try {
      entries = fs.readdirSync(d, { withFileTypes: true });
    } catch (e) {
      return;
    }
    for (const ent of entries) {
      const full = path.join(d, ent.name);
      if (ent.isDirectory()) {
        walk(full);
      } else if (ent.isFile()) {
        found.push(full);
      }
    }
  }
  walk(dir);
  return found;
}

function looksLikeMarkdown(file) {
  try {
    const s = fs.readFileSync(file, 'utf8');
    return s.includes('### ') && s.includes('_[');
  } catch (e) {
    return false;
  }
}

async function waitForDownloadFile(dir, timeoutMs) {
  const before = new Set(listFiles(dir));
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const now = listFiles(dir);
    const fresh = now.filter((f) => !before.has(f));
    // The download dir is a dedicated mkdtemp, so any new file is our download.
    // Wait until its content actually looks like markdown so we don't read a
    // half-written file (Chrome writes data: URIs atomically, but be safe).
    const ready = fresh.find(looksLikeMarkdown);
    if (ready) return ready;
    await new Promise((r) => setTimeout(r, 200));
  }
  return null;
}

function readStatus(page) {
  return page.evaluate(() => {
    const host = Array.from(document.querySelectorAll('div')).find(
      (d) => d.shadowRoot && d.shadowRoot.querySelector('.ttc-status')
    );
    if (!host) return null;
    const s = host.shadowRoot.querySelector('.ttc-status');
    return s ? s.textContent : null;
  });
}

function clickDownloadButton(page) {
  return page.evaluate(() => {
    const host = Array.from(document.querySelectorAll('div')).find(
      (d) => d.shadowRoot && d.shadowRoot.querySelector('[data-action="download"]')
    );
    if (!host) throw new Error('TTC panel host not found');
    const btn = host.shadowRoot.querySelector('[data-action="download"]');
    if (!btn) throw new Error('download button not found');
    if (btn.disabled) throw new Error('download button is disabled (not ready)');
    btn.click();
    return true;
  });
}

// ---------------------------------------------------------------------------
// The test.
// ---------------------------------------------------------------------------

test('extension captures local transcript, shows ready UI, downloads markdown', async (t) => {
  // --- Skip conditions (logged reason, not a failure) ---
  if (!chromium) {
    t.skip('playwright is not installed — run `npm install --save-dev playwright`');
    return;
  }
  if (!CHROMIUM_PATH) {
    t.skip(
      'no Chromium binary found — install Chromium/Chrome or set TTC_CHROMIUM_PATH to a chrome/chromium executable'
    );
    return;
  }

  const userDataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ttc-profile-'));
  const downloadDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ttc-download-'));
  let serverStarted = false;
  let context = null;
  const consoleLogs = [];

  try {
    await startServer();
    serverStarted = true;
  } catch (e) {
    t.skip(`fixture server failed to start: ${String((e && e.message) || e)}`);
    return;
  }

  try {
    const launchOptions = {
      headless: process.env.TTC_HEADED ? false : true,
      executablePath: CHROMIUM_PATH,
      acceptDownloads: true,
      downloadsPath: downloadDir,
      args: [
        `--disable-extensions-except=${REPO_ROOT}`,
        `--load-extension=${REPO_ROOT}`,
        `--host-resolver-rules=MAP ${HOSTNAME} 127.0.0.1`,
        '--ignore-certificate-errors',
      ],
    };

    try {
      context = await chromium.launchPersistentContext(userDataDir, launchOptions);
    } catch (launchErr) {
      const msg = String((launchErr && launchErr.message) || launchErr);
      if (/sandbox/i.test(msg)) {
        // Retry once without the Chromium sandbox (common in containers/CI).
        context = await chromium.launchPersistentContext(userDataDir, {
          ...launchOptions,
          args: [...launchOptions.args, '--no-sandbox'],
        });
      } else {
        throw launchErr;
      }
    }

    const page = context.pages()[0] || (await context.newPage());
    page.on('console', (m) => {
      consoleLogs.push(`[${m.type()}] ${m.text()}`);
      if (m.type() === 'error') {
        // Surface extension/page errors to stderr for triage without failing.
        process.stderr.write(`  [page console] ${m.text()}\n`);
      }
    });

    await page.goto(PAGE_URL, { waitUntil: 'domcontentloaded', timeout: 20000 });

    // Wait up to 30s for the shadow-DOM panel to report ready.
    let status = null;
    try {
      await page.waitForFunction(
        () => {
          const host = Array.from(document.querySelectorAll('div')).find(
            (d) => d.shadowRoot && d.shadowRoot.querySelector('.ttc-status')
          );
          if (!host) return false;
          const s = host.shadowRoot.querySelector('.ttc-status');
          return !!s && /Transcript ready/.test(s.textContent || '');
        },
        { timeout: 30000 }
      );
      status = await readStatus(page);
    } catch (e) {
      status = await readStatus(page);
      const fixtureState = await page.evaluate(() => {
        const meta = window.__fixtureMetadata;
        const err = window.__fixtureError;
        const statusEl = document.getElementById('status');
        return {
          fixtureStatus: statusEl ? statusEl.textContent : null,
          metadataFetched: !!meta,
          fixtureError: err || null,
        };
      });
      assert.fail(
        `floating UI never showed "Transcript ready". ` +
          `status=${JSON.stringify(status)}; fixture=${JSON.stringify(fixtureState)}; ` +
          `consoleLogs=${JSON.stringify(consoleLogs.slice(-20))}`
      );
    }

    assert.match(status, /Transcript ready: \d+ entries/, `unexpected status text: ${status}`);
    const entryCount = Number(/Transcript ready: (\d+) entries/.exec(status)[1]);
    assert.ok(entryCount >= 10, `expected >=10 entries in UI, got ${entryCount}`);

    // Click Download .md and capture the file off disk.
    await clickDownloadButton(page);

    const downloaded = await waitForDownloadFile(downloadDir, 20000);
    assert.ok(downloaded, `no .md file appeared in ${downloadDir} within 20s`);

    const markdown = fs.readFileSync(downloaded, 'utf8');
    const lines = markdown.split('\n');
    const headingLines = lines.filter((l) => l.startsWith('### '));
    const timestampMarkers = lines.filter((l) => l.includes('_['));
    const speakers = new Set(headingLines.map((l) => l.slice('### '.length).trim()));

    assert.ok(headingLines.length >= 10, `expected >=10 "### " headings, got ${headingLines.length}`);
    assert.ok(timestampMarkers.length >= 10, `expected >=10 "_[" timestamp markers, got ${timestampMarkers.length}`);
    assert.ok(speakers.size >= 2, `expected >=2 distinct speakers, got ${speakers.size}: ${[...speakers].join(', ')}`);

    // Non-fatal sanity: the downloaded file should match the UI's reported count.
    t.diagnostic(
      `downloaded ${downloaded}: ${headingLines.length} headings, ${timestampMarkers.length} ` +
        `timestamp markers, ${speakers.size} distinct speakers (UI reported ${entryCount} entries)`
    );
  } finally {
    if (context) {
      try {
        await context.close();
      } catch (e) {
        // ignore
      }
    }
    if (serverStarted) {
      await stopServer();
    }
    for (const d of [userDataDir, downloadDir]) {
      try {
        fs.rmSync(d, { recursive: true, force: true });
      } catch (e) {
        // ignore
      }
    }
  }
});
