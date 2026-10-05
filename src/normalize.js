// src/normalize.js
// Pure module: normalize the three SharePoint/Stream transcript-metadata JSON
// shapes into a single { temporaryDownloadUrl, displayName, languageTag }
// object and rewrite recording transcript download URLs so they return
// plaintext VTT instead of encrypted stream content. No DOM access — loadable
// in Node for tests and in Chrome as a plain script before content.js
// (manifest `js` array ordering).
//
// Shapes handled:
//   1. data.media?.transcripts[]  — Graph API-style; prefer entry where
//      isDefault === true, otherwise fall back to the first entry with a URL.
//   2. data.value[]               — OData collection; same isDefault rule.
//   3. data.temporaryDownloadUrl  — flat response; pass-through.
//
// Returns null when no usable URL exists.

(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    // Browser/Chrome content-script context: attach to a namespaced global so
    // content.js can consume it without needing import maps or bundlers.
    root.TTCNormalize = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  function pickFromArray(arr) {
    if (!Array.isArray(arr) || arr.length === 0) return null;
    // Prefer the default transcript; fall back to the first with a URL.
    const def = arr.find((t) => t && t.isDefault === true && typeof t.temporaryDownloadUrl === "string");
    if (def) return def;
    return arr.find((t) => t && typeof t.temporaryDownloadUrl === "string") || null;
  }

  function toResult(entry) {
    if (!entry) return null;
    return {
      temporaryDownloadUrl: entry.temporaryDownloadUrl,
      displayName: typeof entry.displayName === "string" ? entry.displayName : null,
      languageTag:
        typeof entry.languageTag === "string"
          ? entry.languageTag
          : typeof entry.language === "string"
            ? entry.language
            : null,
    };
  }

  // Detect SharePoint/Stream transcript download URLs that point at the
  // encrypted stream endpoint rather than the plaintext VTT we can parse.
  //
  // For meetings that were both recorded and transcribed, Microsoft stores
  // the transcript as a secondary stream inside the MP4 and the metadata's
  // `temporaryDownloadUrl` may end in `/content` or be a `/streamContent` URL.
  // Hitting those directly returns encrypted bytes; appending `?format=json`
  // to them is also wrong. We rewrite them to `/streamContent?is=1&
  // applymediaedits=false`, which returns the WebVTT file. The regex anchors
  // on `/content` at the end of the path so URLs such as `/contents/` are not
  // accidentally rewritten.
  function isStreamTranscriptUrl(url) {
    if (typeof url !== "string" || !url) return false;
    try {
      const u = new URL(url);
      return u.pathname.endsWith("/content") || u.pathname.includes("/streamContent");
    } catch (e) {
      return false;
    }
  }

  function rewriteStreamTranscriptUrl(url) {
    if (typeof url !== "string" || !url) return url;
    try {
      const u = new URL(url);
      if (u.pathname.endsWith("/content")) {
        u.pathname = u.pathname.slice(0, -"/content".length) + "/streamContent";
      }
      if (u.pathname.includes("/streamContent")) {
        u.search = "?is=1&applymediaedits=false";
        return u.toString();
      }
      return url;
    } catch (e) {
      return url;
    }
  }

  function extractTranscriptUrl(data) {
    if (!data || typeof data !== "object") return null;

    // Shape 1: Graph API — data.media.transcripts[]
    const mediaTranscripts = data.media && Array.isArray(data.media.transcripts) ? data.media.transcripts : null;
    const fromMedia = toResult(pickFromArray(mediaTranscripts));
    if (fromMedia) return fromMedia;

    // Shape 2: OData — data.value[]
    const valueArr = Array.isArray(data.value) ? data.value : null;
    const fromValue = toResult(pickFromArray(valueArr));
    if (fromValue) return fromValue;

    // Shape 3: flat — data.temporaryDownloadUrl
    if (typeof data.temporaryDownloadUrl === "string" && data.temporaryDownloadUrl) {
      return toResult(data);
    }

    return null;
  }

  return { extractTranscriptUrl, isStreamTranscriptUrl, rewriteStreamTranscriptUrl };
});
