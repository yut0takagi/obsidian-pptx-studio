/**
 * The presentation itself: the full-screen surface, and moving through it.
 *
 * What a presenter meets in front of a room is a slide that does not come up,
 * a key that does nothing, a hidden slide shown to the audience, or a show
 * that cannot be left. These drive the real class against a DOM — keys sent
 * the way a browser sends them, fullscreen present and absent — rather than
 * calling its methods, so the wiring is what is being checked.
 */
import { strict as assert } from "node:assert";
import { afterEach, describe, it } from "node:test";
import { installDom } from "../../scripts/dom-shim";

installDom();

import { Presenter } from "../../src/render/Presenter";
import type { SlideTransition } from "../../src/pptx/transition";
import type { Deck, Slide } from "../../src/pptx/types";

function deckOf(hidden: boolean[], transitions: (SlideTransition | null)[] = []): Deck {
	const slides = hidden.map(
		(h, i): Slide => ({
			index: i + 1,
			name: "",
			partPath: `ppt/slides/slide${i + 1}.xml`,
			background: { kind: "solid", color: "#ffffff" },
			shapes: [],
			templateShapes: 0,
			notes: "",
			hidden: h,
			transition: transitions[i] ?? null,
		}),
	);
	return { width: 960, height: 540, slides, title: "Deck" };
}

const fade = (patch: Partial<SlideTransition> = {}): SlideTransition => ({
	effect: "fade",
	duration: 400,
	direction: null,
	orientation: null,
	throughBlack: false,
	advanceAfter: null,
	...patch,
});

const open: Presenter[] = [];
afterEach(() => {
	for (const p of open.splice(0)) p.close();
	delete (window.HTMLElement.prototype as unknown as Record<string, unknown>).requestFullscreen;
	delete (window.HTMLElement.prototype as unknown as Record<string, unknown>).animate;
	delete (document as unknown as Record<string, unknown>).fullscreenElement;
	delete (document as unknown as Record<string, unknown>).exitFullscreen;
});

function start(deck: Deck, startIndex = 0, onClose?: (index: number) => void): Presenter {
	const p = Presenter.open({ deck, startIndex, doc: document, onClose });
	assert.ok(p, "the presentation did not open");
	open.push(p);
	return p;
}

function press(key: string): KeyboardEvent {
	const event = new window.KeyboardEvent("keydown", { key, bubbles: true, cancelable: true });
	document.body.dispatchEvent(event);
	return event;
}

/** The slide number currently on screen, as the renderer labels it. */
function shownSlide(p: Presenter): number | null {
	const slides = p.element.querySelectorAll<HTMLElement>(".pptx-presenter-layer .pptx-slide");
	const last = slides[slides.length - 1];
	return last ? Number(last.dataset.slideIndex) : null;
}

const tick = () => new Promise((resolve) => setTimeout(resolve, 0));

describe("Presenter", () => {
	it("covers the whole window when fullscreen is not available", () => {
		const p = start(deckOf([false, false]));
		assert.equal(p.element.parentElement, document.body);
		assert.ok(p.element.classList.contains("is-windowed"));
	});

	it("asks for fullscreen when it can, and ends the show when fullscreen is left", async () => {
		let fullscreen: Element | null = null;
		const asked: Element[] = [];
		Object.defineProperty(document, "fullscreenElement", {
			configurable: true,
			get: () => fullscreen,
		});
		(window.HTMLElement.prototype as unknown as Record<string, unknown>).requestFullscreen =
			function (this: HTMLElement) {
				asked.push(this);
				fullscreen = this;
				return Promise.resolve();
			};
		let closedAt: number | null = null;
		const p = start(deckOf([false, false]), 0, (index) => (closedAt = index));
		await tick();
		assert.deepEqual(asked, [p.element]);
		assert.ok(!p.element.classList.contains("is-windowed"));

		// Escape in fullscreen belongs to the browser: it leaves fullscreen and
		// the page only hears about it afterwards.
		fullscreen = null;
		document.dispatchEvent(new window.Event("fullscreenchange"));
		assert.equal(p.element.isConnected, false);
		assert.equal(closedAt, 0);
	});

	it("starts on the slide it was opened from", () => {
		const p = start(deckOf([false, false, false]), 1);
		assert.equal(shownSlide(p), 2);
	});

	it("moves on from a hidden slide it was opened on", () => {
		const p = start(deckOf([false, true, false]), 1);
		assert.equal(shownSlide(p), 3);
	});

	it("will not open when every slide is hidden", () => {
		assert.equal(Presenter.open({ deck: deckOf([true, true]), startIndex: 0, doc: document }), null);
	});

	it("pages with the keyboard, skipping hidden slides", () => {
		const p = start(deckOf([false, true, false, false]));
		const event = press("ArrowRight");
		assert.equal(shownSlide(p), 3);
		assert.ok(event.defaultPrevented, "the key also reached whatever is behind the show");
		press(" ");
		assert.equal(shownSlide(p), 4);
		press("PageUp");
		assert.equal(shownSlide(p), 3);
		press("ArrowLeft");
		assert.equal(shownSlide(p), 1);
		press("ArrowLeft");
		assert.equal(shownSlide(p), 1, "went back past the first slide");
	});

	it("advances on a click", () => {
		const p = start(deckOf([false, false]));
		p.element.dispatchEvent(new window.MouseEvent("click", { bubbles: true }));
		assert.equal(shownSlide(p), 2);
	});

	it("ends on a black screen, and the next step leaves", () => {
		let closedAt: number | null = null;
		const p = start(deckOf([false, false, true]), 0, (index) => (closedAt = index));
		press("ArrowRight");
		press("ArrowRight");
		assert.equal(p.slideIndex, null);
		assert.ok(p.element.querySelector(".pptx-presenter-end"));
		assert.equal(closedAt, null);
		press("ArrowRight");
		assert.equal(p.element.isConnected, false);
		assert.equal(closedAt, 1, "should report the last slide shown, not the hidden one after it");
	});

	it("leaves on Escape and says where it got to", () => {
		let closedAt: number | null = null;
		const p = start(deckOf([false, false, false]), 0, (index) => (closedAt = index));
		press("ArrowRight");
		press("Escape");
		assert.equal(p.element.isConnected, false);
		assert.equal(closedAt, 1);
		// Keys after the show belong to Obsidian again.
		assert.equal(press("ArrowRight").defaultPrevented, false);
	});

	it("plays the entering slide's transition going forward, and none going back", async () => {
		const played: Keyframe[][] = [];
		(window.HTMLElement.prototype as unknown as Record<string, unknown>).animate = function (
			frames: Keyframe[],
		) {
			played.push(frames);
			return { finished: Promise.resolve(), cancel() {}, finish() {} };
		};
		const p = start(deckOf([false, false], [null, fade()]));
		press("ArrowRight");
		assert.equal(played.length, 1);
		assert.equal(played[0][0].opacity, 0);
		await tick();
		assert.equal(p.element.querySelectorAll(".pptx-presenter-layer").length, 1, "the old slide was left behind");

		press("ArrowLeft");
		assert.equal(played.length, 1, "going back replayed a transition");
		assert.equal(shownSlide(p), 1);
	});

	it("advances by itself when the slide says to", async () => {
		const p = start(deckOf([false, false], [fade({ effect: "cut", duration: 0, advanceAfter: 5 })]));
		await new Promise((resolve) => setTimeout(resolve, 40));
		assert.equal(shownSlide(p), 2);
	});
});
