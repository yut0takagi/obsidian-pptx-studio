/**
 * Which slide an embed means.
 *
 * `![[deck.pptx#3]]` names a position, and positions move: insert a slide at
 * the front and every note embedding slide 3 now shows what used to be slide
 * 2. `#id=257` names the slide itself — the `p:sldId@id` PowerPoint gives each
 * slide and keeps through reordering — so the embed follows the slide. Both
 * forms are accepted; the number form is what existing notes already contain.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import { installDom } from "../../scripts/dom-shim";

installDom();

import { duplicateSlide, moveSlide } from "../../src/ooxml/slideOps";
import { rebuildDeck, rebuildSlideAt } from "../../src/pptx/parse";
import { parseSlideRef, resolveSlideRef, slideEmbedLink } from "../../src/view/slideRef";
import { miniDeck } from "./miniDeck";

describe("parseSlideRef", () => {
	it("reads the number forms embeds have always used", () => {
		assert.deepEqual(parseSlideRef("3"), { kind: "number", number: 3 });
		assert.deepEqual(parseSlideRef("slide=3"), { kind: "number", number: 3 });
		assert.deepEqual(parseSlideRef("slide3"), { kind: "number", number: 3 });
	});

	it("reads a slide id", () => {
		assert.deepEqual(parseSlideRef("id=257"), { kind: "id", id: 257 });
		assert.deepEqual(parseSlideRef(" ID=300 "), { kind: "id", id: 300 });
	});

	it("means the whole deck for anything else", () => {
		for (const subpath of ["", "0", "abc", "id=", "id=x", "slide=-1"]) {
			assert.equal(parseSlideRef(subpath), undefined, JSON.stringify(subpath));
		}
	});
});

describe("slide ids in the model", () => {
	it("carries each slide's p:sldId@id", () => {
		const deck = rebuildDeck(miniDeck([{ id: 300 }, { id: 257 }]), "Deck");
		assert.deepEqual(
			deck.slides.map((s) => s.slideId),
			[300, 257],
		);
	});

	it("keeps it when one slide is rebuilt after an edit", () => {
		const pkg = miniDeck([{ id: 300 }, { id: 257 }]);
		const deck = rebuildDeck(pkg, "Deck");
		assert.ok(rebuildSlideAt(pkg, deck, 1));
		assert.equal(deck.slides[1].slideId, 257);
	});

	it("gives a duplicated slide an id of its own", () => {
		const pkg = miniDeck([{ id: 300 }, { id: 257 }]);
		duplicateSlide(pkg, "ppt/slides/slide1.xml", 1);
		const ids = rebuildDeck(pkg, "Deck").slides.map((s) => s.slideId);
		assert.equal(new Set(ids).size, ids.length, `ids repeat: ${ids.join(", ")}`);
	});
});

describe("resolveSlideRef", () => {
	it("finds a slide by id wherever it has moved to", () => {
		const pkg = miniDeck([{ id: 256 }, { id: 257 }, { id: 258 }]);
		const before = rebuildDeck(pkg, "Deck");
		const path = before.slides[resolveSlideRef(before, { kind: "id", id: 257 })! - 1].partPath;

		assert.ok(moveSlide(pkg, 1, 0));
		const after = rebuildDeck(pkg, "Deck");
		const n = resolveSlideRef(after, { kind: "id", id: 257 });
		assert.equal(n, 1);
		assert.equal(after.slides[n! - 1].partPath, path, "the id now names a different slide");
		// A number keeps meaning a position.
		assert.equal(resolveSlideRef(after, { kind: "number", number: 2 }), 2);
	});

	it("says when the slide is not in the deck any more", () => {
		const deck = rebuildDeck(miniDeck([{ id: 256 }]), "Deck");
		assert.equal(resolveSlideRef(deck, { kind: "id", id: 999 }), null);
	});

	it("means the whole deck when there is no reference", () => {
		const deck = rebuildDeck(miniDeck([{}]), "Deck");
		assert.equal(resolveSlideRef(deck, undefined), undefined);
	});
});

describe("slideEmbedLink", () => {
	it("embeds a slide by its id", () => {
		assert.equal(
			slideEmbedLink("Talks/deck.pptx", { slideId: 257, index: 2 }),
			"![[Talks/deck.pptx#id=257]]",
		);
	});

	it("falls back to the slide's number when it has no id", () => {
		assert.equal(slideEmbedLink("deck.pptx", { slideId: null, index: 4 }), "![[deck.pptx#4]]");
	});
});
