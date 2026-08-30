import type { PptxPackage } from "../pptx/package";
import { renderSlide } from "../render/renderSlide";
import type { Deck, Slide } from "../pptx/types";

/**
 * Rasterise a slide to PNG.
 *
 * The slide is already plain DOM with inline styles, so it can be dropped into
 * an SVG <foreignObject> and drawn to a canvas. Object URLs do not survive that
 * trip, so every image is re-inlined as a data URL first.
 */
export async function renderSlideToPng(
	deck: Deck,
	slide: Slide,
	pkg: PptxPackage,
	scale: number,
): Promise<ArrayBuffer> {
	const node = renderSlide(deck, slide);
	inlineMedia(node, pkg);

	node.setAttribute("xmlns", "http://www.w3.org/1999/xhtml");
	const html = new XMLSerializer().serializeToString(node);
	const svg =
		`<svg xmlns="http://www.w3.org/2000/svg" width="${deck.width}" height="${deck.height}">` +
		`<foreignObject x="0" y="0" width="${deck.width}" height="${deck.height}">${html}</foreignObject>` +
		`</svg>`;

	const image = await loadImage(
		`data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`,
	);

	const canvas = document.createElement("canvas");
	canvas.width = Math.round(deck.width * scale);
	canvas.height = Math.round(deck.height * scale);
	const ctx = canvas.getContext("2d");
	if (!ctx) throw new Error("Could not get a 2D canvas context.");
	ctx.fillStyle = "#ffffff";
	ctx.fillRect(0, 0, canvas.width, canvas.height);
	ctx.drawImage(image, 0, 0, canvas.width, canvas.height);

	const blob = await new Promise<Blob | null>((resolve) =>
		canvas.toBlob((b) => resolve(b), "image/png"),
	);
	if (!blob) throw new Error("Could not encode the slide as PNG.");
	return blob.arrayBuffer();
}

/** Swap every object URL for a base64 data URL, which survives serialisation. */
function inlineMedia(root: HTMLElement, pkg: PptxPackage): void {
	for (const el of Array.from(root.querySelectorAll<HTMLElement>("[data-media-path]"))) {
		const path = el.dataset.mediaPath;
		if (!path) continue;
		const dataUrl = pkg.mediaDataUrl(path);
		if (!dataUrl) continue;
		if (el instanceof HTMLImageElement) {
			el.src = dataUrl;
		} else {
			el.style.backgroundImage = `url("${dataUrl}")`;
		}
	}
}

function loadImage(src: string): Promise<HTMLImageElement> {
	return new Promise((resolve, reject) => {
		const img = new Image();
		img.onload = () => resolve(img);
		img.onerror = () =>
			reject(new Error("The slide could not be rasterised (unsupported content?)."));
		img.src = src;
	});
}

/** A filename-safe stem for an exported slide, e.g. "deck-slide-03". */
export function pngFileName(deckName: string, slideIndex: number): string {
	const stem = deckName
		.replace(/\.pptx$/i, "")
		.replace(/[\\/:*?"<>|#^[\]]/g, "-")
		.trim();
	return `${stem || "slide"}-slide-${String(slideIndex).padStart(2, "0")}.png`;
}
