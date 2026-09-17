// content.js — runs in the isolated world on *.sharepoint.com at document_idle.
// Owns the bridge from the MAIN-world intercept.js (postMessage intake, strict
// origin guard, one-shot metadata) AND the transcript fetch pipeline: build
// the ?format=json URL, apply MCAS rewrite, race-wait bearer token, fetch in
// the content world with SW fallback, and stash raw text + parsed payload on
// module state. Then converts raw text to Markdown (via src/markdown.js) and
// hands it to the floating UI (src/ui.js) for Copy .md / Download .md.

// Module-level bearer token, updated by TTC_BEARER_TOKEN messages from
// intercept.js. Later waves send it with the transcript fetch.
let bearerToken = null;

// Permanent one-shot flag: the first valid TTC_TRANSCRIPT_METADATA message
// wins for the lifetime of this page load. Subsequent metadata frames are
// ignored rather than overwriting — SharePoint pages can fire the same XHR
// multiple times during hydration/navigation, and we never want to swap
// which URL we captured mid-download.
let metadataHandled = false;

// Wave-3 storage: raw text + parsed payload for the captured transcript.
// `transcriptRawText` is always a string once fetchTextAndStore succeeds;
// `transcriptData` is the JSON.parse()'d object when Content-Type was JSON,
// or null when we had to fall back to VTT text (converter lands in Wave 4).
let transcriptRawText = null;
let transcriptData = null;
let transcriptFormat = null; // "json" | "vtt" | null

// Once the captured transcript text has been converted to Markdown, this holds
// the resulting string (or "" for a transcript with no captions). Drives the
// Copy/Download buttons and the __TTC_STATE__.markdown getter.
let convertedMarkdown = null;

// How long (ms) to wait for a bearer token to arrive via postMessage before
// giving up and issuing the fetch unauthenticated. intercept.js posts the
// token on XHR send() which can race the metadata load; empirically the gap
// is sub-tick, so 500ms is generous while still bounded.
const BEARER_RACE_WAIT_MS = 500;
const BEARER_POLL_INTERVAL_MS = 25;

// Resolve extractTranscriptUrl regardless of load style: src/normalize.js
// either attached itself to window.TTCNormalize (script tag) or was inlined.
const { extractTranscriptUrl } =
  typeof window !== "undefined" && window.TTCNormalize
    ? window.TTCNormalize
    : { extractTranscriptUrl: () => null };

// Converters for the captured raw text. markdown.js requires TTCTime loaded
// first (manifest js ordering guarantees it) and attaches as window.TTCMarkdown.
const { jsonToMarkdown, vttToMarkdown } =
  typeof window !== "undefined" && window.TTCMarkdown
    ? window.TTCMarkdown
    : { jsonToMarkdown: () => "", vttToMarkdown: () => "" };

// Floating shadow-DOM UI. Buttons receive the converted markdown directly from
// ui.js (passed through onCopy/onDownload), so they always act on the current
// capture.
const ui = (typeof window !== "undefined" && window.TTCUI && window.TTCUI.createFloatingUI)
  ? window.TTCUI.createFloatingUI({
      onCopy: (md) => {
        copyToClipboard(md);
      },
      onDownload: (md) => {
        downloadFile(md, "transcript.md");
      },
    })
  : { setStatus: () => {}, setReady: () => {} }; // no-op if ui.js not loaded

// If no metadata arrives and no conversion completes within 30s, surface a hard
// error so the user knows the page isn't a usable transcript tab.
const CAPTURE_TIMEOUT_MS = 30_000;
setTimeout(() => {
  if (convertedMarkdown === null) {
    ui.setStatus(
      "No transcript detected — make sure you're on a meeting Transcript tab, then reload the page",
      "error"
    );
  }
}, CAPTURE_TIMEOUT_MS);

// MCAS (Microsoft Cloud App Security) proxies SharePoint/Teams traffic by
// rewriting hostnames to `<original>.mcas.ms`. The page URL carries the
// rewrite, but metadata JSON still hands us the bare hostname — if we fetch
// the bare URL from an MCAS-proxied page, the request goes around the proxy
// and the tenant's conditional-access policy rejects it. Detect the proxy
// from `window.location` and apply the same suffix to the content hostname.
function applyMcasRewrite(transcriptUrl) {
  try {
    if (!window.location.hostname.endsWith(".mcas.ms")) {
      return transcriptUrl;
    }
    const u = new URL(transcriptUrl);
    if (u.hostname === window.location.hostname) {
      return transcriptUrl; // already rewritten or same-origin
    }
    if (u.hostname.endsWith(".mcas.ms")) {
      return transcriptUrl; // metadata URL is already MCAS-flavored
    }
    u.hostname = u.hostname + ".mcas.ms";
    return u.toString();
  } catch (e) {
    console.warn("[TTC] MCAS rewrite failed, using original URL", e);
    return transcriptUrl;
  }
}

// Append `?format=json` to a SharePoint/Stream download URL. Preserves any
// pre-existing query params; if `format` is already there we leave it alone
// so we don't stack duplicates on re-entrant flows.
function buildJsonUrl(transcriptUrl) {
  const u = new URL(transcriptUrl);
  if (!u.searchParams.has("format")) {
    u.searchParams.set("format", "json");
  }
  return u.toString();
}

function buildAuthHeaders() {
  return bearerToken ? { Authorization: bearerToken } : {};
}

// Race-resilient bearer token: intercept.js posts the token from fetch/XHR
// hooks, and in most flows that precedes the metadata frame. But when the
// XHR carrying the token *is* the metadata call, the token post lands
// synchronously right before our message handler runs. In the rarer
// inverted case (metadata observed first, token XHR dispatched just after),
// we poll briefly before giving up and trying anonymously.
async function waitForBearerToken() {
  if (bearerToken) return bearerToken;
  const deadline = Date.now() + BEARER_RACE_WAIT_MS;
  while (!bearerToken && Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, BEARER_POLL_INTERVAL_MS));
  }
  return bearerToken;
}

function isJsonContentType(contentType) {
  return typeof contentType === "string" && /(^|;|\s)application\/json\b/i.test(contentType);
}

// Issue the transcript fetch from the content-script world. SharePoint pages
// (especially under MCAS or strict CSP) commonly block content-script fetch
// with a Content-Security-Policy `connect-src` violation — that surfaces as
// a rejected promise with TypeError, so ANY throw triggers the SW fallback.
async function fetchFromContent(url) {
  const response = await fetch(url, {
    credentials: "include",
    cache: "no-store",
    headers: buildAuthHeaders(),
  });
  const contentType = response.headers.get("content-type") || "";
  const text = await response.text();
  return { ok: response.ok, status: response.status, text, contentType };
}

// Fallback path: route the same fetch through the MV3 service worker, which
// is not subject to the page's CSP. Kept in lockstep with background.js's
// TTC_FETCH handler — shape is {ok, status, text, contentType}.
async function fetchViaSW(url) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage({ type: "TTC_FETCH", url, credentials: true }, (resp) => {
      const err = chrome.runtime.lastError;
      if (err) {
        return reject(new Error("SW fetch failed: " + err.message));
      }
      if (!resp) {
        return reject(new Error("SW fetch returned empty response"));
      }
      resolve(resp);
    });
  });
}

// Try content-world fetch first, fall back to service worker on any throw.
async function fetchWithFallback(url) {
  try {
    return await fetchFromContent(url);
  } catch (e) {
    console.warn("[TTC] content-world fetch threw (likely CSP), routing via service worker", e);
    return fetchViaSW(url);
  }
}

async function fetchAndStoreTranscript(transcriptUrl) {
  const rewritten = applyMcasRewrite(transcriptUrl);
  const jsonUrl = buildJsonUrl(rewritten);

  await waitForBearerToken();

  let res = await fetchWithFallback(jsonUrl);
  if (!res.ok) {
    throw new Error("transcript fetch failed: HTTP " + res.status);
  }

  if (isJsonContentType(res.contentType)) {
    transcriptRawText = res.text;
    transcriptFormat = "json";
    try {
      transcriptData = JSON.parse(res.text);
    } catch (e) {
      // Server lied about content-type; treat as unparsed and let Wave 4's
      // converter deal with it.
      console.warn("[TTC] JSON content-type but JSON.parse failed", e);
      transcriptData = null;
    }
    console.log("[TTC] transcript JSON stored (%d bytes)", transcriptRawText.length, transcriptRawText);
    finishConversion();
    return;
  }

  // Server returned the VTT payload despite ?format=json — this happens
  // when SharePoint decides the URL token doesn't authorize the JSON
  // variant. Fall back to the original URL (no ?format=json) and stash VTT.
  console.warn("[TTC] non-JSON content-type (%s); falling back to VTT", res.contentType);
  const vttRes = await fetchWithFallback(rewritten);
  if (!vttRes.ok) {
    throw new Error("VTT fallback fetch failed: HTTP " + vttRes.status);
  }
  transcriptRawText = vttRes.text;
  transcriptFormat = "vtt";
  transcriptData = null;
  console.log("[TTC] transcript VTT stored (%d bytes)", transcriptRawText.length, transcriptRawText);

  finishConversion();
}

// Convert the stored raw text to Markdown (JSON or VTT based on the detected
// format), count entries, and surface the result through the UI. No throw:
// garbage input yields "" and the timeout path handles the error messaging.
function finishConversion() {
  let markdown = "";
  if (transcriptFormat === "json") {
    markdown = jsonToMarkdown(transcriptRawText);
  } else if (transcriptFormat === "vtt") {
    markdown = vttToMarkdown(transcriptRawText);
  }
  convertedMarkdown = markdown;

  // Every non-empty speaker block renders as a single "### " heading.
  const entryCount = markdown ? markdown.split("\n").filter((l) => l.startsWith("### ")).length : 0;
  ui.setReady(markdown, entryCount);
}

function handleMetadataMessage(payload) {
  if (metadataHandled) {
    return;
  }
  metadataHandled = true; // one-shot: set before work so re-entry during sync processing is impossible

  const normalized = extractTranscriptUrl(payload);
  if (!normalized || typeof normalized.temporaryDownloadUrl !== "string") {
    console.error("[TTC] metadata received but no usable URL found", payload);
    return;
  }
  console.log("[TTC] transcript metadata captured", normalized);

  // Fire-and-forget: the listener must stay synchronous so we can keep
  // processing subsequent postMessage frames (e.g. late-arriving bearer
  // tokens). Errors land in the console; Wave 4's UI will surface them.
  fetchAndStoreTranscript(normalized.temporaryDownloadUrl).catch((e) => {
    console.error("[TTC] transcript capture failed", e);
  });
}

function handleBearerTokenMessage(payload) {
  if (!payload || typeof payload.token !== "string") {
    console.warn("[TTC] TTC_BEARER_TOKEN received without a string token", payload);
    return;
  }
  bearerToken = payload.token;
  console.log("[TTC] bearer token updated (len=%d)", bearerToken.length);
}

window.addEventListener("message", (event) => {
  // Origin is the only reliable signal from an isolated world: page and
  // content-script `window` are different proxies, so `event.source === window`
  // would reject every legitimate post from intercept.js.
  if (event.origin !== window.location.origin) {
    return;
  }

  const data = event.data;
  if (!data || typeof data !== "object" || typeof data.type !== "string") {
    return;
  }

  switch (data.type) {
    case "TTC_TRANSCRIPT_METADATA":
      // Wire shape: { type, data: <metadata JSON>, url }
      handleMetadataMessage(data.data);
      break;
    case "TTC_BEARER_TOKEN":
      // Wire shape: { type, token } — handler reads .token off the envelope.
      handleBearerTokenMessage(data);
      break;
    default:
      break;
  }
});

// Exported for debugging in DevTools only — later waves may read this via
// chrome.scripting.executeScript if they need to inspect what was captured.
window.__TTC_STATE__ = {
  get bearerToken() {
    return bearerToken;
  },
  get metadataHandled() {
    return metadataHandled;
  },
  get transcriptRawText() {
    return transcriptRawText;
  },
  get transcriptData() {
    return transcriptData;
  },
  get transcriptFormat() {
    return transcriptFormat;
  },
  get markdown() {
    return convertedMarkdown;
  },
};

// ---------------------------------------------------------------------------
// Delivery helpers. No side effects on load — callers invoke these from
// user gestures (Wave 4+ UI) once the pipeline has produced text.

// navigator.clipboard.writeText requires a user gesture; callers must invoke
// this from a click handler, not from a timer or postMessage callback. We
// intentionally do NOT fall back to document.execCommand('copy') — the
// extension has clipboardWrite and writeText is the only non-deprecated
// path in MV3 isolated worlds.
async function copyToClipboard(text) {
  if (typeof text !== "string") {
    return { ok: false, reason: "missing-text" };
  }
  try {
    await navigator.clipboard.writeText(text);
    return { ok: true };
  } catch (err) {
    console.error("[TTC] clipboard write failed:", err);
    return { ok: false, reason: String((err && err.message) || err) };
  }
}

// Route downloads through the service worker. Content scripts can't reach
// chrome.downloads directly, and background.js picks between the data: URI
// path and the offscreen Blob path based on text length. Errors are caught
// and returned as {ok:false, reason} — never throw on a user gesture; the
// UI wave owns surfacing failures via badge/notification.
async function downloadFile(text, filename) {
  if (typeof text !== "string") {
    return { ok: false, reason: "missing-text" };
  }
  if (typeof filename !== "string" || filename.length === 0) {
    return { ok: false, reason: "missing-filename" };
  }
  try {
    const response = await chrome.runtime.sendMessage({
      type: "TTC_DOWNLOAD",
      text,
      filename,
    });
    if (!response || typeof response !== "object") {
      return { ok: false, reason: "no-response" };
    }
    return response;
  } catch (err) {
    console.error("[TTC] TTC_DOWNLOAD sendMessage failed:", err);
    return { ok: false, reason: String((err && err.message) || err) };
  }
}

// DevTools surface so Wave 4 UI wiring (and manual QA via executeScript) can
// reach these without re-parsing this file.
window.__TTC_HELPERS__ = { copyToClipboard, downloadFile };
