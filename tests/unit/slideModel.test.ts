/**
 * What the parser says about a slide as a whole, rather than about its shapes.
 *
 * These are the slide-level facts that decide how a deck is presented and
 * exported: whether a slide is shown at all, and how it enters. They are read
 * through parseDeck on a real (if tiny) package, and again through the
 * single-slide rebuild that runs after an edit, because the second path is the
 * one that quietly forgets things when a field is added to only the first.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import { installDomParser } from "./dom";

installDomParser();

import { rebuildDeck, rebuildSlideAt } from "../../src/pptx/parse";
import { miniDeck } from "./miniDeck";

describe("hidden slides", () => {
	it("reads show=\"0\" as hidden and everything else as shown", () => {
		const pkg = miniDeck([{}, { attrs: 'show="0"' }, { attrs: 'show="1"' }, { attrs: 'show="false"' }]);
		const deck = rebuildDeck(pkg, "Deck");
		assert.deepEqual(
			deck.slides.map((s) => s.hidden),
			[false, true, false, true],
		);
	});

	it("keeps knowing a slide is hidden after the slide alone is rebuilt", () => {
		const pkg = miniDeck([{}, { attrs: 'show="0"' }]);
		const deck = rebuildDeck(pkg, "Deck");
		assert.ok(rebuildSlideAt(pkg, deck, 1));
		assert.equal(deck.slides[1].hidden, true);
	});
});

describe("slide transitions in the model", () => {
	it("carries each slide's transition, and null where there is none", () => {
		const pkg = miniDeck([
			{},
			{ after: '<p:transition spd="fast"><p:push dir="u"/></p:transition>' },
		]);
		const deck = rebuildDeck(pkg, "Deck");
		assert.equal(deck.slides[0].transition, null);
		assert.equal(deck.slides[1].transition?.effect, "push");
		assert.equal(deck.slides[1].transition?.direction, "u");
		assert.equal(deck.slides[1].transition?.duration, 500);
	});
});
