/**
 * The decisions behind presenting: which slides are shown, where a show
 * starts, what a key means, and what a transition looks like as keyframes.
 *
 * They are kept apart from the DOM that acts on them because each is a small
 * rule that is easy to get subtly wrong — a show that starts on a hidden slide,
 * a push that travels the opposite way to PowerPoint's — and easier to pin
 * down as data than by watching an animation.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";

import {
	presentationOrder,
	presenterAction,
	startPosition,
	transitionPlan,
} from "../../src/render/slideshow";
import type { SlideTransition } from "../../src/pptx/transition";
import type { Deck, Slide } from "../../src/pptx/types";

function deckOf(hidden: boolean[]): Deck {
	const slides = hidden.map(
		(h, i): Slide => ({
			index: i + 1,
			name: "",
			partPath: `ppt/slides/slide${i + 1}.xml`,
			slideId: 256 + i,
			background: null,
			shapes: [],
			templateShapes: 0,
			notes: "",
			hidden: h,
			transition: null,
		}),
	);
	return { width: 960, height: 540, slides, title: "Deck" };
}

function transition(patch: Partial<SlideTransition>): SlideTransition {
	return {
		effect: "fade",
		duration: 700,
		direction: null,
		orientation: null,
		throughBlack: false,
		advanceAfter: null,
		...patch,
	};
}

/** A keyframe property's value at the start and at the end. */
function ends(frames: Keyframe[] | null, property: string): [unknown, unknown] {
	assert.ok(frames && frames.length >= 2, `no keyframes for ${property}`);
	return [frames[0][property], frames[frames.length - 1][property]];
}

describe("presentationOrder", () => {
	it("keeps every slide when none is hidden", () => {
		assert.deepEqual(presentationOrder(deckOf([false, false, false])), [0, 1, 2]);
	});

	it("leaves hidden slides out", () => {
		assert.deepEqual(presentationOrder(deckOf([false, true, false, true])), [0, 2]);
	});

	it("is empty when everything is hidden", () => {
		assert.deepEqual(presentationOrder(deckOf([true, true])), []);
	});
});

describe("startPosition", () => {
	it("starts on the slide asked for when it is shown", () => {
		assert.equal(startPosition([0, 2, 3], 2), 1);
	});

	it("moves on to the next shown slide from a hidden one", () => {
		assert.equal(startPosition([0, 2, 3], 1), 1);
	});

	it("falls back to the last shown slide when nothing after it is shown", () => {
		assert.equal(startPosition([0, 2], 3), 1);
	});

	it("has nowhere to start when nothing is shown", () => {
		assert.equal(startPosition([], 0), -1);
	});
});

describe("presenterAction", () => {
	it("advances on the keys a clicker or keyboard sends", () => {
		for (const key of ["ArrowRight", "ArrowDown", "PageDown", " ", "Enter", "n"]) {
			assert.equal(presenterAction({ key }), "next", key);
		}
	});

	it("goes back on the reverse keys", () => {
		for (const key of ["ArrowLeft", "ArrowUp", "PageUp", "Backspace", "p"]) {
			assert.equal(presenterAction({ key }), "previous", key);
		}
	});

	it("jumps to either end, and leaves on Escape", () => {
		assert.equal(presenterAction({ key: "Home" }), "first");
		assert.equal(presenterAction({ key: "End" }), "last");
		assert.equal(presenterAction({ key: "Escape" }), "exit");
	});

	it("leaves shortcuts with a modifier to Obsidian", () => {
		assert.equal(presenterAction({ key: "ArrowRight", metaKey: true }), null);
		assert.equal(presenterAction({ key: "p", ctrlKey: true }), null);
		assert.equal(presenterAction({ key: "q" }), null);
	});
});

describe("transitionPlan", () => {
	it("plays nothing without a transition, for a cut, or when motion is unwelcome", () => {
		assert.equal(transitionPlan(null, { reducedMotion: false }), null);
		assert.equal(transitionPlan(transition({ effect: "cut", duration: 0 }), { reducedMotion: false }), null);
		assert.equal(transitionPlan(transition({ effect: "push", direction: "l" }), { reducedMotion: true }), null);
	});

	it("fades the incoming slide in over the transition's duration", () => {
		const plan = transitionPlan(transition({ effect: "fade", duration: 900 }), { reducedMotion: false });
		assert.equal(plan?.duration, 900);
		assert.deepEqual(ends(plan!.incoming, "opacity"), [0, 1]);
		assert.equal(plan?.outgoing, null);
	});

	it("goes through black by fading the old slide out before the new one in", () => {
		const plan = transitionPlan(transition({ effect: "fade", throughBlack: true }), { reducedMotion: false });
		assert.deepEqual(ends(plan!.outgoing, "opacity"), [1, 0]);
		assert.deepEqual(ends(plan!.incoming, "opacity"), [0, 1]);
		const midway = plan!.incoming.find((frame) => frame.offset === 0.5);
		assert.equal(midway?.opacity, 0, "the new slide is still dark halfway through");
	});

	it("pushes in the direction the effect names", () => {
		// dir is the way the slides travel: "l" moves them left, so the new one
		// arrives from the right.
		const left = transitionPlan(transition({ effect: "push", direction: "l" }), { reducedMotion: false });
		assert.deepEqual(ends(left!.incoming, "transform"), ["translate(100%, 0%)", "translate(0%, 0%)"]);
		assert.deepEqual(ends(left!.outgoing, "transform"), ["translate(0%, 0%)", "translate(-100%, 0%)"]);
		const up = transitionPlan(transition({ effect: "push", direction: "u" }), { reducedMotion: false });
		assert.deepEqual(ends(up!.incoming, "transform"), ["translate(0%, 100%)", "translate(0%, 0%)"]);
	});

	it("covers without moving the old slide, and pulls it away to uncover the new one", () => {
		const cover = transitionPlan(transition({ effect: "cover", direction: "rd" }), { reducedMotion: false });
		assert.deepEqual(ends(cover!.incoming, "transform"), ["translate(-100%, -100%)", "translate(0%, 0%)"]);
		assert.equal(cover?.outgoing, null);

		const pull = transitionPlan(transition({ effect: "pull", direction: "d" }), { reducedMotion: false });
		assert.equal(pull?.outgoingOnTop, true);
		assert.equal(pull?.incoming.length ?? 0, 0);
		assert.deepEqual(ends(pull!.outgoing, "transform"), ["translate(0%, 0%)", "translate(0%, 100%)"]);
	});

	it("wipes the new slide on from the edge it travels away from", () => {
		const left = transitionPlan(transition({ effect: "wipe", direction: "l" }), { reducedMotion: false });
		assert.deepEqual(ends(left!.incoming, "clipPath"), ["inset(0% 0% 0% 100%)", "inset(0% 0% 0% 0%)"]);
		const down = transitionPlan(transition({ effect: "wipe", direction: "d" }), { reducedMotion: false });
		assert.deepEqual(ends(down!.incoming, "clipPath"), ["inset(0% 0% 100% 0%)", "inset(0% 0% 0% 0%)"]);
	});

	it("splits outward by opening the new slide, inward by closing the old one", () => {
		const out = transitionPlan(
			transition({ effect: "split", orientation: "vert", direction: "out" }),
			{ reducedMotion: false },
		);
		assert.deepEqual(ends(out!.incoming, "clipPath"), ["inset(0% 50% 0% 50%)", "inset(0% 0% 0% 0%)"]);

		const inward = transitionPlan(
			transition({ effect: "split", orientation: "horz", direction: "in" }),
			{ reducedMotion: false },
		);
		assert.equal(inward?.outgoingOnTop, true);
		assert.deepEqual(ends(inward!.outgoing, "clipPath"), ["inset(0% 0% 0% 0%)", "inset(50% 0% 50% 0%)"]);
	});

	it("has a drawing for every effect the parser can return", () => {
		for (const effect of ["fade", "push", "wipe", "cover", "pull", "split", "zoom", "circle"] as const) {
			const plan = transitionPlan(transition({ effect, direction: effect === "zoom" ? "in" : "l" }), {
				reducedMotion: false,
			});
			assert.ok(plan, effect);
			assert.ok(plan.incoming.length > 0 || (plan.outgoing?.length ?? 0) > 0, effect);
		}
	});
});
