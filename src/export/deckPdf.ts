/**
 * Turning the whole deck into pages: PDF on disk, or the print dialog.
 *
 * Both go through the same rasteriser the PNG export uses, so a slide looks the
 * same however it leaves the plugin, and a slide that renders in the viewer is
 * a slide that prints.
 */
import type { PptxPackage } from "../pptx/package";
import type { Deck } from "../pptx/types";
import { rasteriseSlide } from "./png";
import { buildPdf, pointSize, rgbFromRgba, type PdfPage } from "./pdf";

/** Rasterise every slide, reporting progress so a long deck can say so. */
async function pages(
	deck: Deck,
	pkg: PptxPackage,
	scale: number,
	onProgress?: (done: number, total: number) => void,
): Promise<PdfPage[]> {
	const out: PdfPage[] = [];
	const size = pointSize(deck.width, deck.height);
	for (const [index, slide] of deck.slides.entries()) {
		const canvas = await rasteriseSlide(deck, slide, pkg, scale);
		const context = canvas.getContext("2d");
		if (!context) throw new Error("Could not get a 2D canvas context.");
		const pixels = context.getImageData(0, 0, canvas.width, canvas.height);
		out.push({
			width: canvas.width,
			height: canvas.height,
			pointWidth: size.width,
			pointHeight: size.height,
			rgb: rgbFromRgba(pixels.data),
		});
		onProgress?.(index + 1, deck.slides.length);
	}
	return out;
}

export async function deckToPdf(
	deck: Deck,
	pkg: PptxPackage,
	scale: number,
	onProgress?: (done: number, total: number) => void,
): Promise<Uint8Array> {
	return buildPdf(await pages(deck, pkg, scale, onProgress));
}

/**
 * Hand the deck to the print dialog.
 *
 * Printing the app's own window would print Obsidian, so the slides go into an
 * off-screen frame of their own with a page box the size of a slide — which is
 * also what makes "Save as PDF" from that dialog come out right.
 */
export async function printDeck(
	deck: Deck,
	pkg: PptxPackage,
	scale: number,
	onProgress?: (done: number, total: number) => void,
): Promise<void> {
	const images: string[] = [];
	for (const [index, slide] of deck.slides.entries()) {
		const canvas = await rasteriseSlide(deck, slide, pkg, scale);
		images.push(canvas.toDataURL("image/png"));
		onProgress?.(index + 1, deck.slides.length);
	}

	const size = pointSize(deck.width, deck.height);
	const frame = document.body.createEl("iframe");
	frame.setCssStyles({
		position: "fixed",
		right: "0",
		bottom: "0",
		width: "0",
		height: "0",
		border: "0",
	});

	const doc = frame.contentDocument;
	const win = frame.contentWindow;
	if (!doc || !win) {
		frame.detach();
		throw new Error("Could not open a print frame.");
	}

	// The frame is its own realm, so Obsidian's additions to Document — createEl
	// and the `win` behind it — are not on this document's prototype. The lint
	// rule that asks for them is right everywhere except here.
	const style = doc.createElement("style");
	style.textContent =
		`@page { size: ${size.width}pt ${size.height}pt; margin: 0; }` +
		"html, body { margin: 0; padding: 0; }" +
		"img { display: block; width: 100%; break-after: page; }" +
		"img:last-child { break-after: auto; }";
	doc.head.appendChild(style);
	for (const src of images) {
		const img = doc.createElement("img");
		img.src = src;
		doc.body.appendChild(img);
	}

	// The dialog is modal, so the frame can only be taken away once it returns.
	try {
		await new Promise((resolve) => window.setTimeout(resolve, 100));
		win.focus();
		win.print();
	} finally {
		frame.detach();
	}
}
