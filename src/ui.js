// src/ui.js
// Floating, shadow-DOM-isolated panel for the capture pipeline. Pure-ish: it
// only touches the DOM it creates and exposes setters — it has no knowledge of
// transcripts, markdown, Chrome APIs, or the page it runs in. Loadable both as
// a Chrome content-script (root.TTCUI) and in Node (module.exports) so the
// fixture harness can exercise it without a browser.
//
// Usage (this file must load before content.js in manifest js ordering):
//   const ui = window.TTCUI.createFloatingUI({ onCopy, onDownload });
//   ui.setStatus("Waiting for transcript…");   // idle / error states
//   ui.setReady(markdown, entryCount);         // enables Copy/Download

(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.TTCUI = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // A small coherent system scoped inside the shadow root: one neutral ramp,
  // one semantic ready/error pair, 8px spacing grid, 12px radius.
  var CSS = [
    ":host { all: initial; }",
    ".ttc-panel {",
    "  box-sizing: border-box;",
    "  width: 264px;",
    "  max-width: calc(100vw - 32px);",
    "  background: #ffffff;",
    "  border: 1px solid #d0d7de;",
    "  border-radius: 12px;",
    "  box-shadow: 0 6px 24px rgba(0, 0, 0, 0.16);",
    "  padding: 12px 14px;",
    "  display: flex;",
    "  flex-direction: column;",
    "  gap: 10px;",
    "  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto,",
    "    Helvetica, Arial, sans-serif;",
    "  font-size: 13px;",
    "  line-height: 1.35;",
    "  color: #1f2328;",
    "  -webkit-font-smoothing: antialiased;",
    "}",
    ".ttc-header { display: flex; align-items: flex-start; justify-content: space-between; gap: 8px; }",
    ".ttc-status {",
    "  flex: 1;",
    "  min-width: 0;",
    "  font-size: 12px;",
    "  color: #57606a;",
    "  min-height: 16px;",
    "  overflow-wrap: anywhere;",
    "}",
    ".ttc-status.is-ready { color: #1a7f37; font-weight: 600; }",
    ".ttc-status.is-error { color: #cf222e; }",
    ".ttc-close {",
    "  appearance: none;",
    "  border: none;",
    "  background: transparent;",
    "  color: #57606a;",
    "  cursor: pointer;",
    "  padding: 2px 6px;",
    "  margin: -2px -6px -2px 0;",
    "  border-radius: 6px;",
    "  font-size: 18px;",
    "  line-height: 1;",
    "  font-weight: 400;",
    "}",
    ".ttc-close:hover { background: #eef1f4; color: #1f2328; }",
    ".ttc-actions { display: flex; gap: 8px; }",
    ".ttc-btn {",
    "  box-sizing: border-box;",
    "  flex: 1;",
    "  appearance: none;",
    "  border: 1px solid #d0d7de;",
    "  background: #f6f8fa;",
    "  color: #1f2328;",
    "  border-radius: 8px;",
    "  padding: 6px 10px;",
    "  font-size: 12px;",
    "  font-weight: 600;",
    "  cursor: pointer;",
    "  transition: background 0.15s ease, border-color 0.15s ease;",
    "}",
    ".ttc-btn:not(:disabled):hover { background: #eef1f4; border-color: #afb8c1; }",
    ".ttc-btn.primary { background: #1f2328; color: #ffffff; border-color: #1f2328; }",
    ".ttc-btn.primary:not(:disabled):hover { background: #32383f; }",
    ".ttc-btn:disabled { opacity: 0.45; cursor: not-allowed; }",
    ""
  ].join("\n");

  // Options: { onCopy(markdown), onDownload(markdown) } button callbacks.
  // Returns { setStatus(text, cls), setReady(markdown, entryCount), host }.
  function createFloatingUI(options) {
    options = options || {};
    var onCopy = typeof options.onCopy === "function" ? options.onCopy : null;
    var onDownload =
      typeof options.onDownload === "function" ? options.onDownload : null;

    // The host div sits on the page root; the shadow root isolates its
    // children. The fixed bottom/right/z-index must be an inline style on the
    // host box itself — an :host rule can't raise the host out of the flow on
    // the document the way this inline style does.
    var host = document.createElement("div");
    host.style.position = "fixed";
    host.style.bottom = "16px";
    host.style.right = "16px";
    host.style.zIndex = "2147483647";

    var shadow = host.attachShadow({ mode: "open" });

    var style = document.createElement("style");
    style.textContent = CSS;
    shadow.appendChild(style);

    // Literal, author-authored markup (no user input) — innerHTML is safe here.
    var rootEl = document.createElement("div");
    rootEl.setAttribute("role", "status");
    rootEl.setAttribute("aria-live", "polite");
    rootEl.innerHTML =
      '<div class="ttc-panel">' +
      '  <div class="ttc-header">' +
      '    <div class="ttc-status">Waiting for transcript…</div>' +
      '    <button type="button" class="ttc-close" data-action="close" aria-label="Close transcript panel">×</button>' +
      "  </div>" +
      '  <div class="ttc-actions">' +
      '    <button type="button" class="ttc-btn" data-action="copy" disabled>Copy .md</button>' +
      '    <button type="button" class="ttc-btn primary" data-action="download" disabled>Download .md</button>' +
      "  </div>" +
      "</div>";
    shadow.appendChild(rootEl);

    document.documentElement.appendChild(host);

    var statusEl = rootEl.querySelector(".ttc-status");
    var copyBtn = rootEl.querySelector('[data-action="copy"]');
    var dlBtn = rootEl.querySelector('[data-action="download"]');
    var closeBtn = rootEl.querySelector('[data-action="close"]');
    var currentMarkdown = null;

    copyBtn.addEventListener("click", function () {
      if (onCopy && currentMarkdown !== null) onCopy(currentMarkdown);
    });
    dlBtn.addEventListener("click", function () {
      if (onDownload && currentMarkdown !== null) onDownload(currentMarkdown);
    });
    closeBtn.addEventListener("click", function () {
      if (host.parentNode) host.parentNode.removeChild(host);
    });

    // Set idle/error status and disable both buttons (not ready).
    function setStatus(text, cls) {
      currentMarkdown = null;
      copyBtn.disabled = true;
      dlBtn.disabled = true;
      statusEl.textContent = text == null ? "" : String(text);
      statusEl.classList.toggle("is-ready", cls === "ready");
      statusEl.classList.toggle("is-error", cls === "error");
    }

    // Transition to ready: store markdown, enable both buttons, show count.
    function setReady(markdown, entryCount) {
      currentMarkdown = typeof markdown === "string" ? markdown : null;
      var count =
        typeof entryCount === "number" && isFinite(entryCount)
          ? Math.max(0, Math.floor(entryCount))
          : 0;
      statusEl.textContent = "Transcript ready: " + count + " entries";
      statusEl.classList.add("is-ready");
      statusEl.classList.remove("is-error");
      copyBtn.disabled = false;
      dlBtn.disabled = false;
    }

    // Transient feedback (e.g. "Copied!") without disturbing the ready state:
    // buttons stay enabled and currentMarkdown is left intact so the user can
    // copy again or download immediately after.
    function notify(text, cls) {
      statusEl.textContent = text == null ? "" : String(text);
      statusEl.classList.toggle("is-ready", cls === "ready");
      statusEl.classList.toggle("is-error", cls === "error");
    }

    return { setStatus: setStatus, setReady: setReady, notify: notify, host: host };
  }

  return { createFloatingUI: createFloatingUI };
});
