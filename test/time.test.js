// test/time.test.js
// Unit tests for src/time.js — parseOffset, formatTimestamp, formatVttTimestamp.

const { test } = require("node:test");
const assert = require("node:assert/strict");
const { parseOffset, formatTimestamp, formatVttTimestamp } = require("../src/time.js");

// ---------------------------------------------------------------------------
// parseOffset — string variants
// ---------------------------------------------------------------------------

test("parseOffset: H:MM:SS", () => {
  assert.equal(parseOffset("0:01:02"), 62);
});

test("parseOffset: HH:MM:SS", () => {
  assert.equal(parseOffset("1:02:03"), 3600 + 120 + 3);
});

test("parseOffset: HH:MM:SS.mmm with fraction", () => {
  assert.equal(parseOffset("1:02:03.5"), 3600 + 120 + 3 + 0.5);
});

test("parseOffset: HH:MM:SS.050 fraction", () => {
  assert.equal(parseOffset("0:00:01.050"), 1 + 0.05);
});

// ---------------------------------------------------------------------------
// parseOffset — numeric variants
// ---------------------------------------------------------------------------

test("parseOffset: ms (>100000)", () => {
  // 120000 ms = 120 s
  assert.equal(parseOffset(120000), 120);
});

test("parseOffset: ticks (>1e10)", () => {
  // 1 second of .NET ticks = 1e7, but must be >1e10 to trigger ticks heuristic
  assert.equal(parseOffset(12_345_678_901), 1234.5678901);
});

test("parseOffset: small number treated as seconds", () => {
  assert.equal(parseOffset(42), 42);
});

// ---------------------------------------------------------------------------
// parseOffset — edge cases
// ---------------------------------------------------------------------------

test("parseOffset: null/undefined returns 0", () => {
  assert.equal(parseOffset(null), 0);
  assert.equal(parseOffset(undefined), 0);
});

test("parseOffset: negative number returns 0", () => {
  assert.equal(parseOffset(-5), 0);
});

test("parseOffset: malformed string returns 0", () => {
  assert.equal(parseOffset("not a time"), 0);
});

test("parseOffset: zero", () => {
  assert.equal(parseOffset(0), 0);
  assert.equal(parseOffset("0:00:00"), 0);
});

// ---------------------------------------------------------------------------
// formatTimestamp
// ---------------------------------------------------------------------------

test("formatTimestamp: 1 second", () => {
  assert.equal(formatTimestamp(1), "0:00:01");
});

test("formatTimestamp: >1 hour", () => {
  assert.equal(formatTimestamp(3723), "1:02:03");
});

test("formatTimestamp: zero", () => {
  assert.equal(formatTimestamp(0), "0:00:00");
});

test("formatTimestamp: rounds down fractional seconds", () => {
  assert.equal(formatTimestamp(1.9), "0:00:01");
});

test("formatTimestamp: negative input coerced to 0", () => {
  assert.equal(formatTimestamp(-10), "0:00:00");
});

test("formatTimestamp: non-number coerced to 0", () => {
  assert.equal(formatTimestamp("hello"), "0:00:00");
});

// ---------------------------------------------------------------------------
// formatVttTimestamp
// ---------------------------------------------------------------------------

test("formatVttTimestamp: zero", () => {
  assert.equal(formatVttTimestamp(0), "00:00:00.000");
});

test("formatVttTimestamp: >1 hour", () => {
  assert.equal(formatVttTimestamp(3723), "01:02:03.000");
});

test("formatVttTimestamp: milliseconds", () => {
  assert.equal(formatVttTimestamp(1.5), "00:00:01.500");
});

test("formatVttTimestamp: single-digit ms zero-padded to 3", () => {
  assert.equal(formatVttTimestamp(0.005), "00:00:00.005");
});

test("formatVttTimestamp: two-digit ms zero-padded", () => {
  assert.equal(formatVttTimestamp(0.05), "00:00:00.050");
});

test("formatVttTimestamp: negative input coerced to 0", () => {
  assert.equal(formatVttTimestamp(-5), "00:00:00.000");
});
