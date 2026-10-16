/**
 * Hidden slides: kept in the deck, visible to whoever edits it, absent from
 * everything an audience sees.
 *
 * A hidden slide is usually a backup — an appendix, a slide for one client but
 * not another — and the failure that matters is it turning up where it was
 * hidden from: in a PDF sent out, or on paper. The other half is the editor:
 * someone working on the deck has to be able to tell which slides those are
 * without opening each one.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import { installDom } from "../../scripts/dom-shim";

installDom();

import { deckToPdf, printedSlides, printDeck } from "../../src/export/deckPdf";
import type { PptxPackage } from "../../src/pptx/package";
import { DeckViewer } from "../../src/render/DeckViewer";
import type { Deck, Slide } from "../../src/pptx/types";

function deckOf(hidden: boolean[]): Deck {
	const slides = hidden.map(
		(h, i): Slide => ({
			index: i + 1,
			name: "",
			partPath: `ppt/slides/slide${i + 1}.xml`,
			slideId: 256 + i,
			background: { kind: "solid", color: "#ffffff" },
			shapes: [],
			templateShapes: 0,
			notes: "",
			hidden: h,
			transition: null,
		}),
	);
	return { width: 960, height: 540, slides, title: "Deck" };
}

/** The observers the viewer uses do not exist under jsdom; they only need to be callable. */
function installObservers(): void {
	class Quiet {
		observe(): void {}
		unobserve(): void {}
		disconnect(): void {}
	}
	const g = globalThis as unknown as Record<string, unknown>;
	g.ResizeObserver = Quiet;
	g.IntersectionObserver = Quiet;
	(window.HTMLElement.prototype as unknown as Record<string, unknown>).scrollIntoView = () => undefined;
}

describe("printed output", () => {
	it("leaves hidden slides out of PDF and print", () => {
		const printed = printedSlides(deckOf([false, true, false, true]));
		assert.deepEqual(
			printed.map((s) => s.index),
			[1, 3],
		);
	});

	it("says so, rather than writing an empty PDF, when every slide is hidden", async () => {
		const pkg = null as unknown as PptxPackage;
		await assert.rejects(deckToPdf(deckOf([true, true]), pkg, 1), /hidden/i);
		await assert.rejects(printDeck(deckOf([true]), pkg, 1), /hidden/i);
	});
});

describe("the thumbnail rail", () => {
	it("marks hidden slides, in a way a screen reader hears too", () => {
		installObservers();
		const host = document.body.createDiv();
		const viewer = new DeckViewer(host, {
			deck: deckOf([false, true, false]),
			pkg: null as unknown as PptxPackage,
			compact: false,
			chrome: "none",
			showThumbnails: true,
			showNotes: false,
			fitMode: "page",
		});
		try {
			const thumbs = Array.from(host.querySelectorAll<HTMLElement>(".pptx-thumb"));
			assert.equal(thumbs.length, 3);
			assert.deepEqual(
				thumbs.map((el) => el.classList.contains("is-hidden-slide")),
				[false, true, false],
			);
			assert.match(thumbs[1].getAttribute("aria-label") ?? "", /hidden/i);
			assert.doesNotMatch(thumbs[0].getAttribute("aria-label") ?? "", /hidden/i);
			assert.ok(thumbs[1].querySelector(".pptx-thumb-hidden"), "no hidden marker beside the number");
		} finally {
			viewer.destroy();
			host.detach();
		}
	});
});
