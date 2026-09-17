// background.js — MV3 service worker for teams-transcript-capture.
// Hosts the TTC_DOWNLOAD pipeline: small payloads go straight to a data: URI;
// large payloads route through the offscreen document for Blob/ObjectURL
// construction (service workers can't call URL.createObjectURL).

const TTC_LARGE_PAYLOAD_THRESHOLD = 2_000_000;
const TTC_OFFSCREEN_URL = "offscreen.html";
const TTC_OFFSCREEN_JUSTIFICATION =
  "Build a Blob/ObjectURL for transcript downloads too large for a data: URI.";
const TTC_OFFSCREEN_CLOSE_TIMEOUT_MS = 5_000;

// Only one offscreen document can exist per profile. Tracks its lifecycle to
// avoid races when multiple downloads could otherwise create it twice.
let offscreenReadyPromise = null;

// In-flight large download requests keyed by a per-request id. The offscreen
// document echoes the requestId back on TTC_BLOB_URL so the SW can correlate.
const pendingBlobRequests = new Map();
let nextBlobRequestId = 1;

async function ensureOffscreenDocument() {
  if (offscreenReadyPromise) {
    return offscreenReadyPromise;
  }
  offscreenReadyPromise = (async () => {
    // chrome.offscreen.hasDocument is available in Chrome 116+; tolerate
    // absence by falling back to createDocument and swallowing the
    // "already exists" error below.
    try {
      if (
        chrome.offscreen.hasDocument &&
        (await chrome.offscreen.hasDocument())
      ) {
        return;
      }
    } catch {
      /* older Chrome — fall through to createDocument */
    }
    try {
      await chrome.offscreen.createDocument({
        url: TTC_OFFSCREEN_URL,
        reasons: ["BLOBS"],
        justification: TTC_OFFSCREEN_JUSTIFICATION,
      });
    } catch (err) {
      if (!String(err && err.message).match(/already exists/i)) {
        throw err;
      }
    }
  })();
  try {
    await offscreenReadyPromise;
  } catch (err) {
    offscreenReadyPromise = null; // allow retry on next call
    throw err;
  }
}

async function closeOffscreenDocumentIfAny() {
  try {
    offscreenReadyPromise = null;
    if (chrome.offscreen.closeDocument) {
      await chrome.offscreen.closeDocument();
    }
  } catch (err) {
    // Never fatal: doc may not exist, or may already have been closed.
    console.warn("[TTC] closeDocument failed (non-fatal):", err && err.message);
  }
}

// Wait for a specific downloadId to reach a terminal state ('complete' or
// 'interrupted'), or time out. Returns the final state string.
function waitForDownloadCompletion(downloadId, timeoutMs) {
  return new Promise((resolve) => {
    let settled = false;
    const timer = setTimeout(() => finish("timeout"), timeoutMs);

    function finish(state) {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      chrome.downloads.onChanged.removeListener(listener);
      resolve(state);
    }

    function listener(delta) {
      if (!delta || delta.id !== downloadId || !delta.state) {
        return;
      }
      const state = delta.state.current;
      if (state === "complete" || state === "interrupted") {
        finish(state);
      }
    }

    chrome.downloads.onChanged.addListener(listener);
  });
}

// Small path: synchronous data: URI, no offscreen involvement.
async function downloadViaDataUri(text, filename) {
  const url =
    "data:text/markdown;charset=utf-8," + encodeURIComponent(text);
  const downloadId = await chrome.downloads.download({
    url,
    filename,
    conflictAction: "uniquify",
  });
  return { ok: true, downloadId, path: "data-uri" };
}

// Large path: request the offscreen document to mint a Blob URL, then wait
// for a matching TTC_BLOB_URL broadcast from offscreen.js before downloading.
async function downloadViaOffscreenBlob(text, filename) {
  await ensureOffscreenDocument();

  const requestId = "ttc-blob-" + nextBlobRequestId++;

  const blobUrlPromise = new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pendingBlobRequests.delete(requestId);
      reject(new Error("offscreen blob timeout"));
    }, TTC_OFFSCREEN_CLOSE_TIMEOUT_MS);

    pendingBlobRequests.set(requestId, {
      resolve: (blobUrl) => {
        clearTimeout(timer);
        pendingBlobRequests.delete(requestId);
        resolve(blobUrl);
      },
      reject: (err) => {
        clearTimeout(timer);
        pendingBlobRequests.delete(requestId);
        reject(err);
      },
      filename,
    });
  });

  try {
    await chrome.runtime.sendMessage({
      type: "TTC_CREATE_BLOB",
      requestId,
      text,
      mimeType: "text/markdown",
    });
  } catch (err) {
    const pending = pendingBlobRequests.get(requestId);
    if (pending) {
      pending.reject(err);
    } else {
      throw err;
    }
  }

  const blobUrl = await blobUrlPromise;

  const downloadId = await chrome.downloads.download({
    url: blobUrl,
    filename,
    conflictAction: "uniquify",
  });

  // Only close the offscreen document once the download has settled — closing
  // earlier would revoke the ObjectURL mid-download. The timeout guarantees
  // the doc doesn't leak even if Chrome never emits a terminal event.
  const finalState = await waitForDownloadCompletion(
    downloadId,
    TTC_OFFSCREEN_CLOSE_TIMEOUT_MS
  );
  await closeOffscreenDocumentIfAny();

  return {
    ok: finalState !== "interrupted",
    downloadId,
    path: "blob",
    finalState,
  };
}

async function handleDownload(message) {
  const { text, filename } = message || {};
  if (typeof text !== "string") {
    return { ok: false, reason: "missing-text" };
  }
  if (typeof filename !== "string" || filename.length === 0) {
    return { ok: false, reason: "missing-filename" };
  }

  try {
    if (text.length < TTC_LARGE_PAYLOAD_THRESHOLD) {
      return await downloadViaDataUri(text, filename);
    }
    return await downloadViaOffscreenBlob(text, filename);
  } catch (err) {
    // Best-effort cleanup in case the offscreen doc came up before failing.
    await closeOffscreenDocumentIfAny();
    console.error("[TTC] TTC_DOWNLOAD failed:", err);
    return { ok: false, reason: String((err && err.message) || err) };
  }
}

// The offscreen document broadcasts TTC_BLOB_URL after TTC_CREATE_BLOB rather
// than replying on the same channel — we use a requestId to correlate because
// multiple in-flight downloads could otherwise race on a single blob slot.
function handleBlobUrl(message) {
  const { requestId, blobUrl, error } = message || {};
  if (typeof requestId !== "string") {
    console.warn("[TTC] TTC_BLOB_URL missing requestId", message);
    return;
  }
  const pending = pendingBlobRequests.get(requestId);
  if (!pending) {
    console.warn("[TTC] TTC_BLOB_URL for unknown requestId", requestId);
    return;
  }
  if (typeof blobUrl === "string") {
    pending.resolve(blobUrl);
  } else {
    pending.reject(new Error(error || "offscreen blob failed"));
  }
}

// Response objects are not structured-cloneable across the message channel,
// so read the body here and ship primitives. Shape matches what
// content.js:fetchViaSW expects — {ok, status, text, contentType}.
async function handleFetch(url) {
  try {
    const response = await fetch(url, { credentials: "include", cache: "no-store" });
    const text = await response.text();
    const contentType = response.headers.get("content-type") || "";
    return { ok: response.ok, status: response.status, text, contentType };
  } catch (e) {
    return { ok: false, status: 0, text: "", contentType: "", error: String((e && e.message) || e) };
  }
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || typeof message.type !== "string") {
    return false;
  }

  switch (message.type) {
    case "TTC_DOWNLOAD":
      // Async path — return true to keep the message channel open.
      handleDownload(message).then(sendResponse, (err) => {
        sendResponse({
          ok: false,
          reason: String((err && err.message) || err),
        });
      });
      return true;

    case "TTC_BLOB_URL":
      handleBlobUrl(message);
      // Ack so offscreen.js doesn't log a message-channel-closed warning.
      sendResponse({ ok: true });
      return false;

    case "TTC_FETCH": {
      const url = message.url;
      if (typeof url !== "string" || url.length === 0) {
        sendResponse({ ok: false, status: 0, text: "", contentType: "", error: "missing-url" });
        return false;
      }
      // Async path — return true to hold the message channel open. The SW is
      // not bound by the page's Content-Security-Policy, so any fetch that
      // the content world sees blocked by connect-src succeeds here.
      handleFetch(url).then(sendResponse);
      return true;
    }

    case "TTC_CREATE_BLOB":
      // This is handled by offscreen.js; if it ever reaches the SW, it means
      // the offscreen document wasn't ready — surface that as an error.
      sendResponse({ ok: false, reason: "offscreen-not-ready" });
      return false;

    default:
      return false;
  }
});
