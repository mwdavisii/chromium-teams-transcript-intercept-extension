// src/time.js
// Pure module: timestamp parsing/formatting helpers for Teams/Stream
// transcript offsets. No DOM access — loadable in Node for tests and in
// Chrome as a plain script before content.js (manifest `js` array ordering).
//
// Offsets seen in the wild come in three shapes:
//   1. ISO-ish duration strings — "H:MM:SS", "HH:MM:SS", "HH:MM:SS.mmm"
//   2. Milliseconds as a number (typically >100000 for any real transcript)
//   3. .NET ticks (100ns units, typically >1e10) from Graph API responses
// parseOffset normalizes all three to seconds (float).

(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    // Browser/Chrome content-script context: attach to a namespaced global so
    // content.js can consume it without needing import maps or bundlers.
    root.TTCTime = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // Matches "H:MM:SS" or "HH:MM:SS" with optional ".mmm" (any digit count).
  // Hours may be any number of digits (e.g. "123:45:56" is valid).
  var OFFSET_RE = /^(\d+):([0-5]?\d):([0-5]?\d)(?:\.(\d+))?$/;

  // Heuristic cutoffs for guessing the unit of a numeric offset.
  var MS_THRESHOLD = 100000; // numbers above this are treated as ms
  var TICKS_THRESHOLD = 1e10; // numbers above this are treated as 100ns ticks

  function parseOffset(offset) {
    if (offset == null) return 0;

    if (typeof offset === "string") {
      var m = OFFSET_RE.exec(offset.trim());
      if (!m) return 0;
      var hours = parseInt(m[1], 10);
      var minutes = parseInt(m[2], 10);
      var seconds = parseInt(m[3], 10);
      var frac = 0;
      if (m[4]) {
        // Treat digits after '.' as a decimal fraction of one second, so
        // ".5" == 500ms and ".050" == 50ms.
        frac = parseInt(m[4], 10) / Math.pow(10, m[4].length);
      }
      return hours * 3600 + minutes * 60 + seconds + frac;
    }

    if (typeof offset === "number" && isFinite(offset)) {
      if (offset < 0) return 0;
      if (offset > TICKS_THRESHOLD) return offset / 1e7; // 100ns ticks → s
      if (offset > MS_THRESHOLD) return offset / 1000; // ms → s
      return offset; // already seconds
    }

    return 0;
  }

  function pad2(n) {
    return n < 10 ? "0" + n : String(n);
  }

  // seconds → "H:MM:SS" (hours unpadded; minutes/seconds zero-padded).
  // Examples: 1 → "0:00:01", 3723 → "1:02:03".
  function formatTimestamp(seconds) {
    if (typeof seconds !== "number" || !isFinite(seconds) || seconds < 0) {
      seconds = 0;
    }
    var total = Math.floor(seconds);
    var hours = Math.floor(total / 3600);
    var minutes = Math.floor((total % 3600) / 60);
    var secs = total % 60;
    return hours + ":" + pad2(minutes) + ":" + pad2(secs);
  }

  // seconds → "HH:MM:SS.mmm" (WebVTT cue timestamp).
  // Hours zero-padded to 2 — VTT consumers accept wider hours, but the spec
  // examples and most tooling stick with 2.
  function formatVttTimestamp(seconds) {
    if (typeof seconds !== "number" || !isFinite(seconds) || seconds < 0) {
      seconds = 0;
    }
    var totalMs = Math.round(seconds * 1000);
    var ms = totalMs % 1000;
    var total = Math.floor(totalMs / 1000);
    var hours = Math.floor(total / 3600);
    var minutes = Math.floor((total % 3600) / 60);
    var secs = total % 60;
    var msStr = ms < 10 ? "00" + ms : ms < 100 ? "0" + ms : String(ms);
    return pad2(hours) + ":" + pad2(minutes) + ":" + pad2(secs) + "." + msStr;
  }

  return { parseOffset, formatTimestamp, formatVttTimestamp };
});
