/**
 * Names out of the file are looked up as names, and nothing more.
 *
 * Effect names, directions, speeds, preset colours, dash styles, chart types,
 * part paths: the plugin turns each into something it can use by looking it up
 * in a table, and the key comes straight out of the deck. A plain object
 * answers "constructor", "toString" or "__proto__" with what it inherits from
 * Object.prototype — a function where a string or a pair of numbers was
 * expected — and what happens next ranges from a wrong colour to a slide show
 * frozen half-way through a transition. Every case here is one of those
 * inherited names, written where a real value would go.
 */
import { strict as assert } from "node:assert";
import { afterEach, describe, it } from "node:test";
import { installDom } from "../../scripts/dom-shim";

installDom();

import { extensionForImageType } from "../../src/edit/clipboardImage";
import { addAnimation, readAnimations } from "../../src/ooxml/animation";
import { parseChartPart } from "../../src/pptx/chart";
import { DEFAULT_THEME, resolveColor } from "../../src/pptx/color";
import { own } from "../../src/pptx/lookup";
import type { PptxPackage } from "../../src/pptx/package";
import { parseStroke } from "../../src/pptx/style";
import { parseSlideTransition, type SlideTransition } from "../../src/pptx/transition";
import type { Deck, Slide } from "../../src/pptx/types";
import { Presenter } from "../../src/render/Presenter";
import { presenterAction, transitionPlan } from "../../src/render/slideshow";
import { miniDeck, miniDeckParts, packageOf } from "./miniDeck";

const INHERITED = ["constructor", "toString", "__proto__", "hasOwnProperty", "valueOf"];
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

function el(xml: string): Element {
	return new DOMParser().parseFromString(xml, "application/xml").documentElement;
}

describe("own", () => {
	it("finds what the table defines, and only that", () => {
		const table: Record<string, number> = { l: 1, r: 2 };
		assert.equal(own(table, "l"), 1);
		for (const key of INHERITED) assert.equal(own(table, key), undefined, key);
		assert.equal(own(table, null), undefined);
		assert.equal(own(table, undefined), undefined);
	});
});

describe("transitions", () => {
	it("plays an inherited direction as the default one, rather than throwing", () => {
		for (const direction of INHERITED) {
			const transition: SlideTransition = {
				effect: "push",
				duration: 500,
				direction,
				orientation: null,
				throughBlack: false,
				advanceAfter: null,
			};
			const plan = transitionPlan(transition, { reducedMotion: false });
			assert.deepEqual(plan?.incoming[0], { transform: "translate(100%, 0%)" }, direction);
		}
	});

	it("reads inherited names in the file as nothing it knows", () => {
		for (const name of INHERITED) {
			const slide = el(
				`<p:sld xmlns:p="${P}"><p:transition spd="${name}" advTm="x"><p:${name}/></p:transition></p:sld>`,
			);
			const t = parseSlideTransition(slide);
			assert.equal(t?.effect, "fade", name);
			assert.equal(t?.duration, 750, `${name}: the speed was not read as the default`);
		}
		const pushed = parseSlideTransition(
			el(`<p:sld xmlns:p="${P}"><p:transition><p:push dir="constructor"/></p:transition></p:sld>`),
		);
		assert.ok(transitionPlan(pushed, { reducedMotion: false }));
	});

	it("is not a key action", () => {
		for (const key of INHERITED) assert.equal(presenterAction({ key }), null, key);
	});
});

describe("a slide show over a transition with an inherited direction", () => {
	afterEach(() => {
		delete (window.HTMLElement.prototype as unknown as Record<string, unknown>).animate;
	});

	it("moves on rather than freezing with both slides on screen", async () => {
		(window.HTMLElement.prototype as unknown as Record<string, unknown>).animate = () => ({
			finished: Promise.resolve(),
			cancel() {},
			finish() {},
		});
		const slide = (index: number, transition: SlideTransition | null): Slide => ({
			index,
			name: "",
			partPath: `ppt/slides/slide${index}.xml`,
			slideId: 255 + index,
			background: { kind: "solid", color: "#ffffff" },
			shapes: [],
			templateShapes: 0,
			notes: "",
			hidden: false,
			transition,
		});
		const deck: Deck = {
			width: 960,
			height: 540,
			title: "Deck",
			slides: [
				slide(1, null),
				slide(2, {
					effect: "push",
					duration: 400,
					direction: "constructor",
					orientation: null,
					throughBlack: false,
					advanceAfter: null,
				}),
			],
		};
		const p = Presenter.open({ deck, startIndex: 0, doc: document });
		assert.ok(p);
		try {
			p.next();
			await new Promise((resolve) => setTimeout(resolve, 0));
			assert.equal(p.slideIndex, 1);
			assert.equal(p.element.querySelectorAll(".pptx-presenter-layer").length, 1);
		} finally {
			p.close();
		}
	});
});

describe("other names read from the file", () => {
	it("gives no colour for an inherited preset or scheme colour name", () => {
		for (const name of INHERITED) {
			const ctx = { theme: DEFAULT_THEME };
			assert.equal(resolveColor(el(`<a:prstClr xmlns:a="${A}" val="${name}"/>`), ctx), null, name);
			assert.equal(resolveColor(el(`<a:schemeClr xmlns:a="${A}" val="${name}"/>`), ctx), null, name);
		}
	});

	it("draws an inherited dash name as a solid line", () => {
		const ctx = { pkg: null as unknown as PptxPackage, partPath: "ppt/slides/slide1.xml", theme: DEFAULT_THEME };
		for (const name of INHERITED) {
			const props = el(
				`<p:spPr xmlns:p="${P}" xmlns:a="${A}"><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill>` +
					`<a:prstDash val="${name}"/></a:ln></p:spPr>`,
			);
			assert.equal(parseStroke(props, ctx)?.style, "solid", name);
		}
	});

	it("reads an animation with inherited preset and trigger names as one it does not know", () => {
		const slide = el(`<p:sld xmlns:p="${P}"><p:cSld><p:spTree/></p:cSld><p:clrMapOvr/></p:sld>`);
		addAnimation(slide, "1", "fade", "click");
		for (const cTn of Array.from(slide.getElementsByTagName("*")).filter((n) => n.localName === "cTn")) {
			if (cTn.getAttribute("presetID")) cTn.setAttribute("presetID", "constructor");
			if (cTn.getAttribute("nodeType") === "clickEffect") cTn.setAttribute("nodeType", "toString");
		}
		const [entry] = readAnimations(slide);
		assert.equal(entry.effect, "other");
		assert.equal(entry.trigger, "click");
	});

	it("skips an element with an inherited name in a chart's plot area", () => {
		const parts = miniDeckParts([{}]);
		parts["ppt/charts/chart1.xml"] =
			`<c:chartSpace xmlns:c="${C}"><c:chart><c:plotArea><c:constructor/>` +
			'<c:barChart><c:barDir val="col"/><c:ser><c:val><c:numLit><c:ptCount val="1"/>' +
			'<c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart>' +
			"</c:plotArea></c:chart></c:chartSpace>";
		const pkg = packageOf(parts);
		const chart = parseChartPart(pkg, "ppt/charts/chart1.xml", {
			pkg,
			partPath: "ppt/slides/slide1.xml",
			theme: DEFAULT_THEME,
		});
		assert.equal(chart.kind, "column");
		assert.equal(chart.groups.length, 1);
	});

	it("has no parts called by inherited names", () => {
		const pkg = miniDeck([{}]);
		for (const name of INHERITED) {
			assert.equal(pkg.has(name), false, name);
			assert.equal(pkg.bytes(name), null, name);
		}
	});

	it("has no file extension for an inherited clipboard type", () => {
		for (const name of INHERITED) assert.equal(extensionForImageType(name), null, name);
	});
});
