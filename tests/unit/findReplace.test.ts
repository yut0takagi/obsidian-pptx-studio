/**
 * Matching and rewriting text that is split across runs.
 *
 * A paragraph is almost never one run — PowerPoint starts a new one at every
 * change of formatting — so the case that matters is the one where the phrase
 * the reader sees straddles two or three of them. These check the plan a
 * replacement makes before any XML is touched.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";

import { findInText, planReplacement, type RunText } from "../../src/edit/findReplace";

const runs = (...texts: string[]): RunText[] => texts.map((text) => ({ text, editable: true }));

/** A run whose text the deck owns — a line break, or a slide-number field. */
const fixed = (text: string): RunText => ({ text, editable: false });

describe("findInText", () => {
	it("finds every occurrence, left to right", () => {
		assert.deepEqual(findInText("a b a b a", "a", true), [
			{ start: 0, length: 1 },
			{ start: 4, length: 1 },
			{ start: 8, length: 1 },
		]);
	});

	it("does not overlap matches with themselves", () => {
		assert.deepEqual(findInText("aaaa", "aa", true), [
			{ start: 0, length: 2 },
			{ start: 2, length: 2 },
		]);
	});

	it("honours case when asked, and ignores it otherwise", () => {
		assert.equal(findInText("Deck deck DECK", "deck", true).length, 1);
		assert.equal(findInText("Deck deck DECK", "deck", false).length, 3);
	});

	it("is literal, not a regular expression", () => {
		assert.deepEqual(findInText("a.c abc", "a.c", true), [{ start: 0, length: 3 }]);
	});

	it("finds nothing for an empty query", () => {
		assert.deepEqual(findInText("anything", "", false), []);
	});
});

describe("planReplacement", () => {
	it("rewrites a match inside one run", () => {
		const input = runs("the quarterly report");
		const plan = planReplacement(input, findInText(input[0].text, "quarterly", true), "annual");
		assert.deepEqual(plan, ["the annual report"]);
	});

	it("replaces every occurrence in a run without disturbing what is between them", () => {
		const input = runs("cat and cat and cat");
		const plan = planReplacement(input, findInText(input[0].text, "cat", true), "dog");
		assert.deepEqual(plan, ["dog and dog and dog"]);
	});

	it("puts a match spanning two runs into the first, and trims the second", () => {
		// "quarter" + "ly" — the second run is bold, which is why it exists.
		const input = runs("the quarter", "ly report");
		const plan = planReplacement(input, [{ start: 4, length: 9 }], "annual");
		assert.deepEqual(plan, ["the annual", " report"]);
	});

	it("empties a run the match swallows whole", () => {
		const input = runs("quar", "ter", "ly done");
		const plan = planReplacement(input, [{ start: 0, length: 9 }], "annual");
		assert.deepEqual(plan, ["annual", "", " done"]);
	});

	it("leaves a match alone when it runs across something it may not rewrite", () => {
		// A line break sits between the two halves: rewriting across it would
		// silently delete the break.
		const input = [{ text: "quar", editable: true }, fixed("\n"), { text: "terly", editable: true }];
		assert.equal(planReplacement(input, [{ start: 0, length: 10 }], "annual"), null);
	});

	it("still replaces the matches that do not touch it", () => {
		const input = [{ text: "cat ", editable: true }, fixed("\n"), { text: "cat", editable: true }];
		const text = input.map((r) => r.text).join("");
		const plan = planReplacement(input, findInText(text, "cat", true), "dog");
		assert.deepEqual(plan, ["dog ", "\n", "dog"]);
	});

	it("reports no plan when nothing would change", () => {
		assert.equal(planReplacement(runs("nothing here"), [], "x"), null);
	});

	it("handles a replacement that contains the search text without looping", () => {
		const input = runs("a a a");
		const plan = planReplacement(input, findInText("a a a", "a", true), "aa");
		assert.deepEqual(plan, ["aa aa aa"]);
	});

	it("deletes a phrase when the replacement is empty", () => {
		const input = runs("keep ", "drop ", "keep");
		const plan = planReplacement(input, [{ start: 5, length: 5 }], "");
		assert.deepEqual(plan, ["keep ", "", "keep"]);
	});

	it("replaces a match that ends exactly on a run boundary", () => {
		const input = runs("abc", "def");
		const plan = planReplacement(input, [{ start: 0, length: 3 }], "X");
		assert.deepEqual(plan, ["X", "def"]);
	});

	it("replaces a match that starts exactly on a run boundary", () => {
		const input = runs("abc", "def");
		const plan = planReplacement(input, [{ start: 3, length: 3 }], "X");
		assert.deepEqual(plan, ["abc", "X"]);
	});
});
