/**
 * Hyperlinks in slide text.
 *
 * A link's target comes out of the deck's relationships, and a deck is a file
 * someone else may have made. Only web and mail links become links; anything
 * else — a script URL, a local file, an app's own scheme — is shown as the text
 * it is, without anything to click. During a slide show a click is how the
 * presenter moves on, so there no slide text is a link at all.
 */
import { strict as assert } from "node:assert";
import { afterEach, describe, it } from "node:test";
import { installDom } from "../../scripts/dom-shim";

installDom();

import { Presenter } from "../../src/render/Presenter";
import { renderSlide } from "../../src/render/renderSlide";
import type { Deck, Run, Shape, Slide } from "../../src/pptx/types";

function run(text: string, link: string | null): Run {
	return {
		text,
		source: null,
		size: 24,
		bold: false,
		italic: false,
		underline: false,
		strike: false,
		color: "#000000",
		font: "Arial",
		baseline: 0,
		spacing: 0,
		link,
		highlight: null,
	};
}

function deckWithLinks(links: (string | null)[]): Deck {
	const shape: Shape = {
		kind: "shape",
		id: "2",
		name: "Text",
		frame: { x: 0, y: 0, w: 960, h: 540, rot: 0, flipH: false, flipV: false },
		hidden: false,
		source: null,
		sourcePart: "ppt/slides/slide1.xml",
		geom: "rect",
		fill: null,
		stroke: null,
		placeholder: null,
		text: {
			source: null,
			sourcePart: "ppt/slides/slide1.xml",
			anchor: "top",
			insets: [0, 0, 0, 0],
			wrap: true,
			fontScale: 1,
			lineSpaceReduction: 0,
			vertical: "horz",
			paragraphs: links.map((link, i) => ({
				source: null,
				level: 0,
				align: "left",
				bullet: null,
				marginLeft: 0,
				indent: 0,
				spaceBefore: 0,
				spaceAfter: 0,
				lineSpacing: null,
				runs: [run(`link ${i}`, link)],
			})),
		},
	};
	const slide = (index: number): Slide => ({
		index,
		name: "",
		partPath: `ppt/slides/slide${index}.xml`,
		slideId: 255 + index,
		background: { kind: "solid", color: "#ffffff" },
		shapes: index === 1 ? [shape] : [],
		templateShapes: 0,
		notes: "",
		hidden: false,
		transition: null,
	});
	return { width: 960, height: 540, slides: [slide(1), slide(2)], title: "Deck" };
}

/** The rendered run for each paragraph, in order. */
function runs(root: Element): HTMLElement[] {
	return Array.from(root.querySelectorAll<HTMLElement>("[data-run]"));
}

describe("links in the deck view", () => {
	it("keeps web and mail links as links", () => {
		const links = ["https://example.com/a", "http://example.com/b", "mailto:someone@example.com"];
		const rendered = runs(renderSlide(deckWithLinks(links), deckWithLinks(links).slides[0]));
		assert.deepEqual(
			rendered.map((el) => el.tagName),
			["A", "A", "A"],
		);
		assert.deepEqual(
			rendered.map((el) => el.getAttribute("href")),
			links,
		);
		assert.equal(rendered[0].getAttribute("rel"), "noopener");
	});

	it("does not care how the scheme is capitalised", () => {
		const deck = deckWithLinks(["HTTPS://EXAMPLE.COM/"]);
		const [link] = runs(renderSlide(deck, deck.slides[0]));
		assert.equal(link.tagName, "A");
		assert.equal(link.getAttribute("href"), "https://example.com/");
	});

	it("shows any other target as plain text", () => {
		const links = [
			"javascript:alert(1)",
			"JavaScript:void(0)",
			" javascript:alert(1)",
			"java\tscript:alert(1)",
			"vbscript:msgbox(1)",
			"file:///etc/hosts",
			"data:text/html,<b>hi</b>",
			"obsidian://open?vault=example",
			"app://obsidian.md/index.html",
			"//example.com/",
			"ppt/slides/slide2.xml",
		];
		const deck = deckWithLinks(links);
		const root = renderSlide(deck, deck.slides[0]);
		assert.equal(root.querySelectorAll("a").length, 0, "some target became a link");
		assert.equal(root.querySelectorAll("[href]").length, 0);
		assert.deepEqual(
			runs(root).map((el) => el.textContent),
			links.map((_, i) => `link ${i}`),
			"the text of the run went missing",
		);
	});
});

describe("links during a slide show", () => {
	const open: Presenter[] = [];
	afterEach(() => {
		for (const p of open.splice(0)) p.close();
	});

	it("draws link text without a link, so a click cannot follow it", () => {
		const p = Presenter.open({ deck: deckWithLinks(["https://example.com/"]), startIndex: 0, doc: document });
		assert.ok(p);
		open.push(p);
		assert.equal(p.element.querySelectorAll("a").length, 0);
		assert.equal(runs(p.element)[0]?.textContent, "link 0");
	});

	it("treats a click on link text as moving on", () => {
		const p = Presenter.open({ deck: deckWithLinks(["https://example.com/"]), startIndex: 0, doc: document });
		assert.ok(p);
		open.push(p);
		const event = new window.MouseEvent("click", { bubbles: true, cancelable: true });
		runs(p.element)[0].dispatchEvent(event);
		assert.equal(p.slideIndex, 1);
		assert.ok(event.defaultPrevented, "the click was left to do whatever it does by default");
	});
});
