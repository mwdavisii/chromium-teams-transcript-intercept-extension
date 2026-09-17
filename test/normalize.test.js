// test/normalize.test.js
// Unit tests for src/normalize.js — extractTranscriptUrl.

const { test } = require("node:test");
const assert = require("node:assert/strict");
const { extractTranscriptUrl } = require("../src/normalize.js");

// ---------------------------------------------------------------------------
// Shape 1 — data.media.transcripts[]
// ---------------------------------------------------------------------------

test("extractTranscriptUrl: prefers isDefault=true in media.transcripts", () => {
  const data = {
    media: {
      transcripts: [
        { temporaryDownloadUrl: "http://first", isDefault: false },
        { temporaryDownloadUrl: "http://second", isDefault: true },
        { temporaryDownloadUrl: "http://third", isDefault: false },
      ],
    },
  };
  const result = extractTranscriptUrl(data);
  assert.equal(result.temporaryDownloadUrl, "http://second");
});

test("extractTranscriptUrl: falls back to first with URL when no isDefault", () => {
  const data = {
    media: {
      transcripts: [
        { temporaryDownloadUrl: "http://first" },
        { temporaryDownloadUrl: "http://second" },
      ],
    },
  };
  const result = extractTranscriptUrl(data);
  assert.equal(result.temporaryDownloadUrl, "http://first");
});

test("extractTranscriptUrl: skips entries without URL in media.transcripts", () => {
  const data = {
    media: {
      transcripts: [
        { isDefault: true },
        { temporaryDownloadUrl: "http://has-url" },
      ],
    },
  };
  const result = extractTranscriptUrl(data);
  assert.equal(result.temporaryDownloadUrl, "http://has-url");
});

test("extractTranscriptUrl: null when media.transcripts is empty", () => {
  assert.equal(extractTranscriptUrl({ media: { transcripts: [] } }), null);
});

// ---------------------------------------------------------------------------
// Shape 2 — data.value[]
// ---------------------------------------------------------------------------

test("extractTranscriptUrl: prefers isDefault=true in value[]", () => {
  const data = {
    value: [
      { temporaryDownloadUrl: "http://a", isDefault: false },
      { temporaryDownloadUrl: "http://b", isDefault: true },
    ],
  };
  const result = extractTranscriptUrl(data);
  assert.equal(result.temporaryDownloadUrl, "http://b");
});

test("extractTranscriptUrl: falls back to first with URL in value[]", () => {
  const data = {
    value: [
      { temporaryDownloadUrl: "http://fallback" },
    ],
  };
  const result = extractTranscriptUrl(data);
  assert.equal(result.temporaryDownloadUrl, "http://fallback");
});

test("extractTranscriptUrl: null when value is empty", () => {
  assert.equal(extractTranscriptUrl({ value: [] }), null);
});

// ---------------------------------------------------------------------------
// Shape 3 — flat data.temporaryDownloadUrl
// ---------------------------------------------------------------------------

test("extractTranscriptUrl: flat temporaryDownloadUrl", () => {
  const data = {
    temporaryDownloadUrl: "http://flat",
    displayName: "Transcript",
    languageTag: "en-US",
  };
  const result = extractTranscriptUrl(data);
  assert.equal(result.temporaryDownloadUrl, "http://flat");
  assert.equal(result.displayName, "Transcript");
  assert.equal(result.languageTag, "en-US");
});

// ---------------------------------------------------------------------------
// Precedence and edge cases
// ---------------------------------------------------------------------------

test("extractTranscriptUrl: media.transcripts wins over value[]", () => {
  const data = {
    media: { transcripts: [{ temporaryDownloadUrl: "http://media", isDefault: true }] },
    value: [{ temporaryDownloadUrl: "http://value", isDefault: true }],
  };
  const result = extractTranscriptUrl(data);
  assert.equal(result.temporaryDownloadUrl, "http://media");
});

test("extractTranscriptUrl: value[] wins over flat", () => {
  const data = {
    value: [{ temporaryDownloadUrl: "http://value" }],
    temporaryDownloadUrl: "http://flat",
  };
  const result = extractTranscriptUrl(data);
  assert.equal(result.temporaryDownloadUrl, "http://value");
});

test("extractTranscriptUrl: null on missing/empty data", () => {
  assert.equal(extractTranscriptUrl(null), null);
  assert.equal(extractTranscriptUrl(undefined), null);
  assert.equal(extractTranscriptUrl({}), null);
});

test("extractTranscriptUrl: falls back to language field when languageTag absent", () => {
  const data = {
    temporaryDownloadUrl: "http://flat",
    language: "fr-FR",
  };
  const result = extractTranscriptUrl(data);
  assert.equal(result.languageTag, "fr-FR");
});

test("extractTranscriptUrl: displayName and languageTag null when absent", () => {
  const data = {
    temporaryDownloadUrl: "http://bare",
  };
  const result = extractTranscriptUrl(data);
  assert.equal(result.temporaryDownloadUrl, "http://bare");
  assert.equal(result.displayName, null);
  assert.equal(result.languageTag, null);
});
