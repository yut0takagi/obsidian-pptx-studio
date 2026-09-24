/**
 * The line height of a shape that holds no text.
 *
 * Decks are full of shapes whose `a:txBody` is empty — a rule four pixels tall,
 * a coloured band behind a heading. The parser still gives such a body one run,
 * sized from the style chain or, failing that, the 18pt default. Rendered
 * literally that is a 60px line box inside a 4px shape: nothing is drawn, but
 * the height shoves the shape's contents out of place.
 *
 * The deck that surfaced this had 129 of them across 34 slides.
 *
 * The paragraph itself stays: the editor finds somewhere to put the caret by
 * looking for `.pptx-para`, so removing it would make an empty shape
 * impossible to type into.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import { installDom } from "../../scripts/dom-shim";

installDom();

import { renderSlide } from "../../src/render/renderSlide";
import type { Deck, Paragraph, Run, Shape, Slide, TextBody } from "../../src/pptx/types";

function run(text: string, size = 60): Run {
	return {
		text,
		source: null,
		size,
		bold: false,
		italic: false,
		underline: false,
		strike: false,
		color: "#000000",
		font: "Arial",
		baseline: 0,
		spacing: 0,
		link: null,
		highlight: null,
	} as Run;
}

function para(...runs: Run[]): Paragraph {
	return {
		source: null,
		level: 0,
		align: "left",
		bullet: null,
		marginLeft: 0,
		indent: 0,
		spaceBefore: 12,
		spaceAfter: 12,
		lineSpacing: null,
		runs,
	};
}

function textBody(...paragraphs: Paragraph[]): TextBody {
	return {
		source: null,
		sourcePart: "ppt/slides/slide1.xml",
		anchor: "top",
		insets: [0, 0, 0, 0],
		wrap: true,
		fontScale: 1,
		lineSpaceReduction: 0,
		paragraphs,
		vertical: "horz",
	};
}

function shape(text: TextBody): Shape {
	return {
		kind: "shape",
		id: "1",
		name: "Rectangle",
		// Four pixels tall: the decoration this is all about.
		frame: { x: 0, y: 0, w: 200, h: 4, rot: 0, flipH: false, flipV: false },
		hidden: false,
		source: null,
		sourcePart: "ppt/slides/slide1.xml",
		geom: "rect",
		fill: null,
		stroke: null,
		text,
		placeholder: null,
	} as Shape;
}

function render(body: TextBody): HTMLElement[] {
	const slide: Slide = {
		index: 1,
		name: "Slide 1",
		partPath: "ppt/slides/slide1.xml",
		background: null,
		shapes: [shape(body)],
		templateShapes: 0,
		notes: "",
	};
	const deck: Deck = { width: 960, height: 540, slides: [slide], title: "deck" };

	return Array.from(renderSlide(deck, slide).querySelectorAll<HTMLElement>(".pptx-para"));
}

describe("a text body with nothing in it", () => {
	it("collapses the line of a body whose only run is empty", () => {
		const [p] = render(textBody(para(run(""))));

		assert.equal(p.style.lineHeight, "0");
	});

	it("drops the paragraph margins with it", () => {
		// 12px above and below a line that shows nothing is still 24px of shove.
		const [p] = render(textBody(para(run(""))));

		assert.equal(p.style.marginTop, "0px");
		assert.equal(p.style.marginBottom, "0px");
	});

	it("collapses every paragraph when none of them has text", () => {
		const paragraphs = render(textBody(para(run("")), para(run("")), para(run(""))));

		assert.equal(paragraphs.length, 3);
		assert.deepEqual(
			paragraphs.map((p) => p.style.lineHeight),
			["0", "0", "0"],
		);
	});

	it("keeps the paragraph element, so the shape can still be typed into", () => {
		const paragraphs = render(textBody(para(run(""))));

		assert.equal(paragraphs.length, 1);
	});

	it("leaves a body with text alone", () => {
		const [p] = render(textBody(para(run("Heading"))));

		assert.notEqual(p.style.lineHeight, "0");
		assert.equal(p.style.marginTop, "12px");
	});

	it("leaves a blank line inside real text alone", () => {
		// This is the spacing case: the empty paragraph is there to push the
		// second line down, so its height has to survive.
		const paragraphs = render(textBody(para(run("First")), para(run("")), para(run("Second"))));

		assert.equal(paragraphs.length, 3);
		for (const p of paragraphs) assert.notEqual(p.style.lineHeight, "0");
	});
});
