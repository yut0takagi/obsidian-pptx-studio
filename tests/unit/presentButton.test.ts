/**
 * Starting a show from the deck view.
 *
 * The palette commands reach the presenter through the view, which needs
 * Obsidian to construct; the button under the slide is the part of that path
 * that can be driven here, and the one most people will find first.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import { installDom } from "../../scripts/dom-shim";

installDom();

import type { PptxPackage } from "../../src/pptx/package";
import { DeckViewer, type DeckViewerOptions } from "../../src/render/DeckViewer";
import type { Deck, Slide } from "../../src/pptx/types";

function deckOf(count: number): Deck {
	const slides = Array.from(
		{ length: count },
		(_, i): Slide => ({
			index: i + 1,
			name: "",
			partPath: `ppt/slides/slide${i + 1}.xml`,
			slideId: 256 + i,
			background: { kind: "solid", color: "#ffffff" },
			shapes: [],
			templateShapes: 0,
			notes: "",
			hidden: false,
			transition: null,
		}),
	);
	return { width: 960, height: 540, slides, title: "Deck" };
}

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

function viewer(options: Partial<DeckViewerOptions>): { host: HTMLElement; viewer: DeckViewer } {
	installObservers();
	const host = document.body.createDiv();
	return {
		host,
		viewer: new DeckViewer(host, {
			deck: deckOf(3),
			pkg: null as unknown as PptxPackage,
			compact: false,
			chrome: "status",
			showThumbnails: false,
			showNotes: false,
			fitMode: "page",
			...options,
		}),
	};
}

describe("the present button", () => {
	it("starts the show from the slide on screen", () => {
		const started: number[] = [];
		const { host, viewer: v } = viewer({ onPresent: (index) => started.push(index) });
		try {
			v.go(1);
			const button = host.querySelector<HTMLElement>('.pptx-statusbar [data-icon="presentation"]');
			assert.ok(button, "no present button in the status bar");
			button.click();
			assert.deepEqual(started, [1]);
		} finally {
			v.destroy();
			host.detach();
		}
	});

	it("is not offered where nothing can start a show", () => {
		const { host, viewer: v } = viewer({});
		try {
			assert.equal(host.querySelector('[data-icon="presentation"]'), null);
		} finally {
			v.destroy();
			host.detach();
		}
	});
});
