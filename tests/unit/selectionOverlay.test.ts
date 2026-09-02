/**
 * The furniture drawn over a slide.
 *
 * It never reaches the file, so nothing here can corrupt a deck — but it is the
 * whole of the editing affordance, and a handle drawn at the wrong size or a
 * rotate knob offered on a multi-selection is a bug a user meets immediately.
 */
import { strict as assert } from "node:assert";
import { beforeEach, describe, it } from "node:test";
import { installDom } from "../../scripts/dom-shim";

installDom();

import { SelectionOverlay, type OverlayItem } from "../../src/edit/SelectionOverlay";
import type { Frame, Shape } from "../../src/pptx/types";

function frame(x: number, y: number, w: number, h: number, rot = 0): Frame {
	return { x, y, w, h, rot, flipH: false, flipV: false };
}

function item(id: string, f: Frame): OverlayItem {
	const shape: Shape = {
		kind: "shape",
		id,
		name: id,
		frame: f,
		hidden: false,
		source: null,
		sourcePart: "ppt/slides/slide1.xml",
		geom: "rect",
		fill: null,
		stroke: null,
		text: null,
		placeholder: null,
	};
	return { shape, el: document.createElement("div") };
}

let slideEl: HTMLElement;
let overlay: SelectionOverlay;
let scale = 1;

beforeEach(() => {
	document.body.innerHTML = "";
	scale = 1;
	slideEl = document.createElement("div");
	slideEl.style.width = "960px";
	slideEl.style.height = "540px";
	document.body.appendChild(slideEl);
	overlay = new SelectionOverlay(() => scale);
	overlay.setSlide(slideEl);
});

const box = () => slideEl.querySelector<HTMLElement>(".pptx-selection:not(.is-member)");
const handles = () => Array.from(slideEl.querySelectorAll<HTMLElement>(".pptx-handle"));
const members = () => Array.from(slideEl.querySelectorAll(".pptx-selection.is-member"));

describe("sync", () => {
	it("draws nothing until something is selected", () => {
		overlay.sync([], false);
		assert.equal(box(), null);
		assert.equal(handles().length, 0);
	});

	it("puts a box with all eight handles around a single shape", () => {
		overlay.sync([item("a", frame(10, 20, 30, 40))], false);
		const el = box();
		assert.ok(el);
		assert.equal(el.style.left, "10px");
		assert.equal(el.style.top, "20px");
		assert.equal(el.style.width, "30px");
		assert.equal(el.style.height, "40px");
		assert.deepEqual(
			handles().map((h) => h.dataset.handle),
			["nw", "n", "ne", "e", "se", "s", "sw", "w"],
		);
	});

	it("offers rotation for one shape only", () => {
		overlay.sync([item("a", frame(0, 0, 10, 10))], false);
		assert.ok(slideEl.querySelector(".pptx-rotate"));

		overlay.sync([item("a", frame(0, 0, 10, 10)), item("b", frame(50, 50, 10, 10))], false);
		assert.equal(slideEl.querySelector(".pptx-rotate"), null);
	});

	it("outlines each member and bounds the lot with one box", () => {
		overlay.sync([item("a", frame(10, 20, 30, 40)), item("b", frame(100, 0, 10, 10))], false);
		assert.equal(members().length, 2);
		const el = box();
		assert.ok(el);
		assert.equal(el.style.left, "10px");
		assert.equal(el.style.top, "0px");
		assert.equal(el.style.width, "100px");
		assert.equal(el.style.height, "60px");
	});

	it("says which state the box is in with its border", () => {
		overlay.sync([item("a", frame(0, 0, 10, 10))], false);
		assert.match(box()?.style.outline ?? "", /solid/);
		overlay.sync([item("a", frame(0, 0, 10, 10))], true);
		assert.match(box()?.style.outline ?? "", /dashed/);
	});

	it("carries a shape's rotation onto its box", () => {
		overlay.sync([item("a", frame(0, 0, 10, 10, 30))], false);
		assert.equal(box()?.style.transform, "rotate(30deg)");
	});

	it("counter-scales the handles, so they stay the same size on screen", () => {
		scale = 0.5;
		overlay.sync([item("a", frame(0, 0, 100, 100))], false);
		assert.equal(handles()[0].style.width, "18px");

		scale = 2;
		overlay.sync([item("a", frame(0, 0, 100, 100))], false);
		assert.equal(handles()[0].style.width, "4.5px");
	});

	it("redraws rather than accumulating", () => {
		overlay.sync([item("a", frame(0, 0, 10, 10))], false);
		overlay.sync([item("a", frame(0, 0, 10, 10))], false);
		assert.equal(slideEl.querySelectorAll(".pptx-selection").length, 1);
		assert.equal(handles().length, 8);
	});

	it("reuses its layers instead of rebuilding them on every sync", () => {
		overlay.sync([item("a", frame(0, 0, 10, 10))], false);
		const layer = slideEl.querySelector(".pptx-overlay");
		overlay.sync([item("a", frame(0, 0, 20, 20))], false);
		assert.equal(slideEl.querySelector(".pptx-overlay"), layer);
	});
});

describe("preview", () => {
	it("moves the box and its handles without touching the model", () => {
		const one = item("a", frame(0, 0, 10, 10));
		overlay.sync([one], false);
		overlay.preview([one], new Map([["a", frame(100, 200, 40, 40)]]));

		assert.equal(box()?.style.left, "100px");
		assert.equal(box()?.style.top, "200px");
		// The south-east handle sits at the far corner, less half its size.
		const se = handles().find((h) => h.dataset.handle === "se");
		assert.equal(se?.style.left, "35.5px");
		// The shape's own frame is untouched: the commit happens on pointer up.
		assert.deepEqual({ x: one.shape.frame.x, y: one.shape.frame.y }, { x: 0, y: 0 });
	});

	it("bounds a multi-shape drag by the frames it was handed", () => {
		const a = item("a", frame(0, 0, 10, 10));
		const b = item("b", frame(50, 50, 10, 10));
		overlay.sync([a, b], false);
		overlay.preview(
			[a, b],
			new Map([
				["a", frame(10, 10, 10, 10)],
				["b", frame(60, 60, 10, 10)],
			]),
		);
		assert.equal(box()?.style.left, "10px");
		assert.equal(box()?.style.width, "60px");
	});
});

describe("guides and marquee", () => {
	it("draws one line per matched guide, replacing the previous set", () => {
		overlay.sync([item("a", frame(0, 0, 10, 10))], false);
		overlay.showGuides({ xs: [10, 20], ys: [30] });
		assert.equal(slideEl.querySelectorAll(".pptx-guide").length, 3);
		overlay.showGuides({ xs: [10], ys: [] });
		assert.equal(slideEl.querySelectorAll(".pptx-guide").length, 1);
	});

	it("keeps one marquee and moves it, rather than leaving a trail", () => {
		overlay.showMarquee({ x: 0, y: 0, w: 10, h: 10 });
		overlay.showMarquee({ x: 5, y: 5, w: 20, h: 20 });
		const all = slideEl.querySelectorAll<HTMLElement>(".pptx-marquee");
		assert.equal(all.length, 1);
		assert.equal(all[0].style.left, "5px");
		assert.equal(all[0].style.width, "20px");
	});

	it("clears the guides and the marquee when a gesture ends, keeping the selection", () => {
		overlay.sync([item("a", frame(0, 0, 10, 10))], false);
		overlay.showGuides({ xs: [10], ys: [10] });
		overlay.showMarquee({ x: 0, y: 0, w: 10, h: 10 });
		overlay.endGesture();
		assert.equal(slideEl.querySelectorAll(".pptx-guide").length, 0);
		assert.equal(slideEl.querySelectorAll(".pptx-marquee").length, 0);
		assert.ok(box(), "the selection box survives the gesture");
	});

	it("draws a marquee again after one ended", () => {
		overlay.showMarquee({ x: 0, y: 0, w: 10, h: 10 });
		overlay.endGesture();
		overlay.showMarquee({ x: 1, y: 1, w: 2, h: 2 });
		assert.equal(slideEl.querySelectorAll(".pptx-marquee").length, 1);
	});
});

describe("clear", () => {
	it("takes both layers off the slide", () => {
		overlay.sync([item("a", frame(0, 0, 10, 10))], false);
		overlay.clear();
		assert.equal(slideEl.querySelector(".pptx-overlay"), null);
		assert.equal(slideEl.querySelector(".pptx-guides"), null);
	});

	it("leaves nothing behind when the slide changes", () => {
		overlay.sync([item("a", frame(0, 0, 10, 10))], false);
		overlay.setSlide(null);
		assert.equal(slideEl.querySelectorAll("div").length, 0);
		overlay.sync([item("a", frame(0, 0, 10, 10))], false);
		assert.equal(slideEl.querySelector(".pptx-selection"), null);
	});
});
