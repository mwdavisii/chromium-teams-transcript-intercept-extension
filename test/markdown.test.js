// test/markdown.test.js
// Unit tests for src/markdown.js — jsonToMarkdown and vttToMarkdown.

const { test } = require("node:test");
const assert = require("node:assert/strict");
const { jsonToMarkdown, vttToMarkdown } = require("../src/markdown.js");

// ---------------------------------------------------------------------------
// jsonToMarkdown
// ---------------------------------------------------------------------------

test("jsonToMarkdown: empty entries returns empty string", () => {
  const data = { entries: [] };
  assert.equal(jsonToMarkdown(data), "");
});

test("jsonToMarkdown: single entry", () => {
  const data = {
    entries: [
      { speakerDisplayName: "Alice", startOffset: "0:00:01", endOffset: "0:00:04", text: "Hello world." },
    ],
  };
  const md = jsonToMarkdown(data);
  assert.ok(md.includes("### Alice"));
  assert.ok(md.includes("_[0:00:01 → 0:00:04]_"));
  assert.ok(md.includes("Hello world."));
});

test("jsonToMarkdown: consecutive same-speaker entries are merged", () => {
  const data = {
    entries: [
      { speakerDisplayName: "Alice", startOffset: "0:00:01", endOffset: "0:00:04", text: "First." },
      { speakerDisplayName: "Alice", startOffset: "0:00:05", endOffset: "0:00:08", text: "Second." },
    ],
  };
  const md = jsonToMarkdown(data);
  const lines = md.split("\n");
  const aliceLines = lines.filter((l) => l.includes("Alice"));
  assert.equal(aliceLines.length, 1); // only one heading
  assert.ok(md.includes("First. Second."));
  assert.ok(md.includes("_[0:00:01 → 0:00:08]_"));
});

test("jsonToMarkdown: missing speaker becomes Unknown", () => {
  const data = {
    entries: [{ startOffset: "0:00:01", endOffset: "0:00:02", text: "No name." }],
  };
  const md = jsonToMarkdown(data);
  assert.ok(md.includes("### Unknown"));
});

test("jsonToMarkdown: empty speaker becomes Unknown", () => {
  const data = {
    entries: [{ speakerDisplayName: "", startOffset: "0:00:01", endOffset: "0:00:02", text: "Empty name." }],
  };
  const md = jsonToMarkdown(data);
  assert.ok(md.includes("### Unknown"));
});

test("jsonToMarkdown: missing text becomes empty and is skipped", () => {
  const data = {
    entries: [
      { speakerDisplayName: "Alice", startOffset: "0:00:01", endOffset: "0:00:02" },
    ],
  };
  assert.equal(jsonToMarkdown(data), "");
});

test("jsonToMarkdown: special characters are escaped", () => {
  const data = {
    entries: [
      {
        speakerDisplayName: "Alice",
        startOffset: "0:00:01",
        endOffset: "0:00:02",
        text: "a*b_c[d#e>f",
      },
    ],
  };
  const md = jsonToMarkdown(data);
  assert.ok(md.includes("a\\*b\\_c\\[d\\#e\\>f"));
});

test("jsonToMarkdown: numeric offsets with realistic ms values", () => {
  // Use values >= 100000 to avoid the <100000 seconds heuristic.
  const data = {
    entries: [
      { speakerDisplayName: "Bob", startOffset: 120000, endOffset: 125000, text: "Numeric offsets." },
    ],
  };
  const md = jsonToMarkdown(data);
  assert.ok(md.includes("### Bob"));
  assert.ok(md.includes("_[0:02:00 → 0:02:05]_"));
});

test("jsonToMarkdown: 1000-entry stress test", () => {
  const entries = [];
  for (let i = 0; i < 1000; i++) {
    entries.push({
      speakerDisplayName: "Speaker",
      startOffset: i,
      endOffset: i + 1,
      text: `word${i}`,
    });
  }
  const md = jsonToMarkdown({ entries });
  // All 1000 entries from the same speaker should merge into ONE segment.
  assert.ok(md.includes("### Speaker"));
  const headingCount = md.split("\n").filter((l) => l.startsWith("### ")).length;
  assert.equal(headingCount, 1);
  assert.ok(md.includes("word0"));
  assert.ok(md.includes("word999"));
});

test("jsonToMarkdown: accepts JSON string", () => {
  const data = JSON.stringify({
    entries: [{ speakerDisplayName: "Alice", startOffset: "0:00:01", endOffset: "0:00:02", text: "From string." }],
  });
  const md = jsonToMarkdown(data);
  assert.ok(md.includes("From string."));
});

test("jsonToMarkdown: returns empty string on invalid JSON string", () => {
  assert.equal(jsonToMarkdown("not json"), "");
});

// ---------------------------------------------------------------------------
// vttToMarkdown
// ---------------------------------------------------------------------------

test("vttToMarkdown: WEBVTT header is stripped", () => {
  const vtt = `WEBVTT\n\n00:00:01.000 --> 00:00:04.000\n<v Alice>Hello.</v>`;
  const md = vttToMarkdown(vtt);
  assert.ok(md.includes("### Alice"));
  assert.ok(md.includes("Hello."));
  assert.ok(!md.includes("WEBVTT"));
});

test("vttToMarkdown: <v Speaker> extraction", () => {
  const vtt = `WEBVTT\n\n00:00:01.000 --> 00:00:03.000\n<v Bob>How are you?</v>`;
  const md = vttToMarkdown(vtt);
  assert.ok(md.includes("### Bob"));
  assert.ok(md.includes("How are you?"));
});

test("vttToMarkdown: cue timestamp parsing", () => {
  const vtt = `WEBVTT\n\n00:00:05.000 --> 00:00:08.500\n<v Alice>Timed.</v>`;
  const md = vttToMarkdown(vtt);
  assert.ok(md.includes("_[0:00:05 → 0:00:08]_"));
});

test("vttToMarkdown: speaker grouping", () => {
  const vtt = `WEBVTT\n\n00:00:01.000 --> 00:00:03.000\n<v Alice>First.</v>\n\n00:00:04.000 --> 00:00:06.000\n<v Alice>Second.</v>`;
  const md = vttToMarkdown(vtt);
  const headingCount = md.split("\n").filter((l) => l.startsWith("### ")).length;
  assert.equal(headingCount, 1);
  assert.ok(md.includes("First. Second."));
});

test("vttToMarkdown: missing speaker becomes Unknown", () => {
  const vtt = `WEBVTT\n\n00:00:01.000 --> 00:00:02.000\nNo voice tag.`;
  const md = vttToMarkdown(vtt);
  assert.ok(md.includes("### Unknown"));
  assert.ok(md.includes("No voice tag."));
});

test("vttToMarkdown: CRLF normalization", () => {
  const vtt = "WEBVTT\r\n\r\n00:00:01.000 --> 00:00:02.000\r\n<v Alice>CRLF.</v>";
  const md = vttToMarkdown(vtt);
  assert.ok(md.includes("### Alice"));
});

test("vttToMarkdown: cue settings ignored", () => {
  const vtt = `WEBVTT\n\n00:00:01.000 --> 00:00:02.000 position:50%\n<v Alice>Settings.</v>`;
  const md = vttToMarkdown(vtt);
  assert.ok(md.includes("Settings."));
});

test("vttToMarkdown: multi-line payload", () => {
  const vtt = `WEBVTT\n\n00:00:01.000 --> 00:00:02.000\n<v Alice>Line one\nLine two</v>`;
  const md = vttToMarkdown(vtt);
  assert.ok(md.includes("Line one Line two"));
});

test("vttToMarkdown: returns empty string for non-string input", () => {
  assert.equal(vttToMarkdown(null), "");
  assert.equal(vttToMarkdown(42), "");
});
