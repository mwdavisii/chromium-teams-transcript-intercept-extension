// offscreen.js — loaded by offscreen.html.
// Owns DOM-only APIs the service worker can't reach (URL.createObjectURL).
// Contract with background.js:
//   Receives  TTC_CREATE_BLOB  {type, requestId, text, mimeType}
//   Broadcasts TTC_BLOB_URL    {type, requestId, blobUrl}  (or {error})
// The requestId lets the SW match Blob URLs to the in-flight TTC_DOWNLOAD.

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || message.type !== "TTC_CREATE_BLOB") {
    return false;
  }

  const { requestId, text, mimeType } = message;

  // Ack the TTC_CREATE_BLOB immediately so the SW's chrome.runtime.sendMessage
  // resolves — the actual blob URL travels back as a separate broadcast.
  sendResponse({ ok: true, received: true });

  (async () => {
    try {
      if (typeof requestId !== "string") {
        throw new Error("missing requestId");
      }
      if (typeof text !== "string") {
        throw new Error("missing text");
      }
      const blob = new Blob([text], {
        type: typeof mimeType === "string" ? mimeType : "text/markdown",
      });
      const blobUrl = URL.createObjectURL(blob);
      await chrome.runtime.sendMessage({
        type: "TTC_BLOB_URL",
        requestId,
        blobUrl,
      });
    } catch (err) {
      // Never let an exception strand the SW's pending promise — always
      // broadcast *something* for this requestId.
      try {
        await chrome.runtime.sendMessage({
          type: "TTC_BLOB_URL",
          requestId: typeof requestId === "string" ? requestId : "unknown",
          error: String((err && err.message) || err),
        });
      } catch (inner) {
        // SW may already have terminated — nothing useful to do here.
        console.error("[TTC offscreen] failed to report blob error:", inner);
      }
    }
  })();

  // We already acked synchronously above, so the message channel can close.
  return false;
});
