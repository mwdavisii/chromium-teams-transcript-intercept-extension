// src/markdown.js
// Pure module: convert SharePoint/Teams transcript payloads (JSON or WebVTT)
// into speaker-grouped Markdown. No DOM access — loadable in Node for tests
// and in Chrome as a plain script before content.js.
//
// Output shape for both converters:
//
//   ### Alice
//
//   _[0:00:01 → 0:00:04]_ First sentence. Second sentence.
//
//   ### Bob
//
//   _[0:00:05 → 0:00:08]_ Reply here.
//
// Consecutive entries from the same speaker are merged into a single bullet
// with their texts joined by a space; the merged timestamp spans from the
// first entry's start to the last entry's end.

(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory(require("./time.js"));
  } else {
    // Browser/Chrome content-script context: expects TTCTime to already be on
    // the global (manifest `js` ordering must load src/time.js first).
    root.TTCMarkdown = factory(root.TTCTime);
  }
})(typeof self !== "undefined" ? self : this, function (TTCTime) {
  "use strict";

  var parseOffset = TTCTime.parseOffset;
  var formatTimestamp = TTCTime.formatTimestamp;

  // VTT cue-timestamp line. Allows cue settings after the end time
  // (e.g. "00:00:01.000 --> 00:00:04.000 position:50%"), which we ignore.
  var VTT_TIMING_RE = /^(\S+)\s+-->\s+(\S+)(?:\s+.*)?$/;

  // Opening voice tag "<v Speaker>" — Speaker captured up to the first '>'.
  var VTT_VOICE_OPEN_RE = /^<v ([^>]+)>/;

  // Markdown special characters that must be backslash-escaped in body text.
  // Per the task spec: *, _, [, #, >.
  var MD_SPECIAL_RE = /([*_[#>])/g;

  function escapeMarkdown(text) {
    if (typeof text !== "string") return "";
    return text.replace(MD_SPECIAL_RE, "\\$1");
  }

  // Meeting/transcript title fallback: missing, null, or whitespace-only
  // titles render as "Transcript" rather than a blank heading.
  function normalizeTitle(title) {
    return typeof title === "string" && title.trim() ? title : "Transcript";
  }

  // Prepend an H1 heading to the rendered body. A transcript with no captions
  // still yields "" (no dangling heading) so callers can distinguish "no
  // content" from "content with a default title".
  function withTitle(title, body) {
    if (typeof body !== "string" || body.length === 0) return "";
    return "# " + normalizeTitle(title) + "\n\n" + body;
  }

  function normalizeSpaces(s) {
    return s.replace(/\s+/g, " ").trim();
  }

  // Group consecutive entries with the same speaker. Each input entry is
  // { speaker, start (seconds), end (seconds), text }. Returns segments.
  function groupBySpeaker(entries) {
    var segments = [];
    for (var i = 0; i < entries.length; i++) {
      var e = entries[i];
      if (!e || typeof e.text !== "string" || e.text.length === 0) continue;
      var last = segments.length > 0 ? segments[segments.length - 1] : null;
      if (last && last.speaker === e.speaker) {
        last.text += " " + e.text;
        if (e.start < last.start) last.start = e.start;
        if (e.end > last.end) last.end = e.end;
      } else {
        segments.push({
          speaker: e.speaker,
          start: e.start,
          end: e.end,
          text: e.text,
        });
      }
    }
    return segments;
  }

  function renderSegments(segments) {
    var out = [];
    for (var i = 0; i < segments.length; i++) {
      var s = segments[i];
      out.push("### " + s.speaker);
      out.push("");
      out.push(
        "_[" +
          formatTimestamp(s.start) +
          " → " +
          formatTimestamp(s.end) +
          "]_ " +
          escapeMarkdown(s.text)
      );
      out.push("");
    }
    return out.join("\n");
  }

  // Convert SharePoint/Stream transcript JSON text to Markdown.
  // Accepts a JSON string or a pre-parsed object. Optional `title` is
  // prepended as an H1 heading (falling back to "Transcript"). Returns "" on
  // garbage or when there are no captions.
  function jsonToMarkdown(jsonText, title) {
    var data = jsonText;
    if (typeof jsonText === "string") {
      try {
        data = JSON.parse(jsonText);
      } catch (e) {
        return "";
      }
    }
    if (!data || typeof data !== "object") return "";
    var rawEntries = Array.isArray(data.entries) ? data.entries : [];
    var entries = [];
    for (var i = 0; i < rawEntries.length; i++) {
      var r = rawEntries[i];
      if (!r || typeof r !== "object") continue;
      entries.push({
        speaker:
          typeof r.speakerDisplayName === "string" && r.speakerDisplayName
            ? r.speakerDisplayName
            : "Unknown",
        start: parseOffset(r.startOffset),
        end: parseOffset(r.endOffset),
        text: normalizeSpaces(typeof r.text === "string" ? r.text : ""),
      });
    }
    return withTitle(title, renderSegments(groupBySpeaker(entries)));
  }

  // Convert WebVTT (with `<v Speaker>` voice tags) to Markdown.
  // Cue blocks are split by blank lines. Each block is:
  //   [optional cue identifier line]
  //   HH:MM:SS.mmm --> HH:MM:SS.mmm [cue settings]
  //   <v Speaker>payload</v>   (payload may span multiple lines)
  // Optional `title` is prepended as an H1 heading (falling back to
  // "Transcript"). Returns "" on unparseable input.
  function vttToMarkdown(vttText, title) {
    if (typeof vttText !== "string") return "";

    // Strip BOM, normalize newlines, drop the WEBVTT header (everything up
    // to the first blank line if the file starts with WEBVTT).
    var body = vttText.replace(/^﻿/, "").replace(/\r\n?/g, "\n");
    if (/^WEBVTT[^\n]*\n/.test(body)) {
      var blank = body.indexOf("\n\n");
      body = blank === -1 ? "" : body.slice(blank + 2);
    }

    var blocks = body.split(/\n{2,}/);
    var entries = [];
    for (var b = 0; b < blocks.length; b++) {
      var lines = blocks[b].split("\n");
      var li = 0;
      // Skip a leading cue-identifier line (numeric index or free-form label).
      if (li < lines.length && !VTT_TIMING_RE.test(lines[li])) li++;
      if (li >= lines.length) continue;
      var timing = VTT_TIMING_RE.exec(lines[li]);
      if (!timing) continue;
      li++;
      var start = parseOffset(timing[1]);
      var end = parseOffset(timing[2]);

      var payload = normalizeSpaces(lines.slice(li).join(" "));
      if (!payload) continue;

      var speaker = "Unknown";
      var text = payload;
      var vm = VTT_VOICE_OPEN_RE.exec(payload);
      if (vm) {
        speaker = vm[1].trim() || "Unknown";
        text = payload.slice(vm[0].length).replace(/<\/v>\s*$/, "");
      }
      text = normalizeSpaces(text);
      if (!text) continue;

      entries.push({ speaker: speaker, start: start, end: end, text: text });
    }
    return withTitle(title, renderSegments(groupBySpeaker(entries)));
  }

  return { jsonToMarkdown, vttToMarkdown };
});
