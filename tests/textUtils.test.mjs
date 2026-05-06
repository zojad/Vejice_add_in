import test from "node:test";
import assert from "node:assert/strict";

import {
  extractTokenBoundaryMetadata,
  normalizeParagraphForEquality,
  onlyCommasChanged,
} from "../src/logic/engine/textUtils.js";

test("extractTokenBoundaryMetadata preserves leading and trailing quotes", () => {
  const result = extractTokenBoundaryMetadata('"beseda,"');

  assert.equal(result.leadingBoundary, '"');
  assert.equal(result.trailingBoundary, ',"');
  assert.equal(result.cleanText, "beseda");
});

test("normalizeParagraphForEquality collapses spacing around commas", () => {
  const result = normalizeParagraphForEquality("To  je  primer ,ki dela.");

  assert.equal(result, "To je primer, ki dela.");
});

test("onlyCommasChanged ignores comma-only edits", () => {
  assert.equal(onlyCommasChanged("To je primer", "To je, primer"), true);
  assert.equal(onlyCommasChanged("To je primer", "To ni primer"), false);
});
