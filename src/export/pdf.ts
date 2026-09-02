/**
 * Writing a deck out as a PDF, without a PDF library.
 *
 * The plugin's whole point is that a deck opens with nothing installed
 * alongside it, and shipping a renderer's worth of dependency to produce a
 * fixed-layout page would undo that. A PDF whose every page is one image is a
 * small enough corner of the format to write directly: a catalogue, a page
 * tree, and per page a content stream that stretches one image XObject over
 * the whole media box.
 *
 * The pixels go in losslessly, deflated — the same compressor the .pptx itself
 * is unpacked with. Slides are mostly flat colour, which deflates well; a deck
 * of full-bleed photographs will produce a large file, and that is the honest
 * trade for not re-encoding someone's images.
 */
import { zlibSync } from "fflate";

/** CSS pixels are 96 to the inch; PDF works in 72nds. */
const PX_TO_POINT = 72 / 96;

export interface PdfPage {
	/** Raster size in pixels. */
	width: number;
	height: number;
	/** Page size in points. */
	pointWidth: number;
	pointHeight: number;
	/** `width * height * 3` bytes: 8-bit RGB, no alpha. */
	rgb: Uint8Array;
}

const encoder = new TextEncoder();

function ascii(text: string): Uint8Array {
	return encoder.encode(text);
}

function concat(chunks: Uint8Array[]): Uint8Array {
	const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
	const out = new Uint8Array(total);
	let at = 0;
	for (const chunk of chunks) {
		out.set(chunk, at);
		at += chunk.length;
	}
	return out;
}

/** Page size in points for a slide measured in CSS pixels. */
export function pointSize(pxWidth: number, pxHeight: number): { width: number; height: number } {
	return {
		width: Math.round(pxWidth * PX_TO_POINT * 100) / 100,
		height: Math.round(pxHeight * PX_TO_POINT * 100) / 100,
	};
}

/** Drop the alpha channel from canvas pixels, which PDF has no use for. */
export function rgbFromRgba(rgba: Uint8ClampedArray | Uint8Array): Uint8Array {
	const out = new Uint8Array((rgba.length / 4) * 3);
	for (let i = 0, o = 0; i < rgba.length; i += 4) {
		out[o++] = rgba[i];
		out[o++] = rgba[i + 1];
		out[o++] = rgba[i + 2];
	}
	return out;
}

/**
 * Assemble the PDF.
 *
 * Objects are numbered as they are written and their byte offsets recorded on
 * the way past, because the cross-reference table at the end has to name where
 * each one starts — which is the one part of the format that cannot be fixed up
 * afterwards without rewriting it.
 */
export function buildPdf(pages: PdfPage[]): Uint8Array {
	const chunks: Uint8Array[] = [];
	const offsets: number[] = [];
	let length = 0;

	const push = (chunk: Uint8Array): void => {
		chunks.push(chunk);
		length += chunk.length;
	};

	/** Write one indirect object, remembering where it began. */
	const object = (id: number, body: Uint8Array[]): void => {
		offsets[id] = length;
		push(ascii(`${id} 0 obj\n`));
		for (const part of body) push(part);
		push(ascii("\nendobj\n"));
	};

	push(ascii("%PDF-1.4\n"));
	// A comment of high bytes marks the file as binary for anything that still
	// transfers files in text mode.
	push(new Uint8Array([0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a]));

	// 1 is the catalogue and 2 the page tree; each page then takes three.
	const pageId = (i: number): number => 3 + i * 3;
	const contentId = (i: number): number => 4 + i * 3;
	const imageId = (i: number): number => 5 + i * 3;

	object(1, [ascii("<< /Type /Catalog /Pages 2 0 R >>")]);
	object(2, [
		ascii(
			`<< /Type /Pages /Count ${pages.length} /Kids [${pages
				.map((_, i) => `${pageId(i)} 0 R`)
				.join(" ")}] >>`,
		),
	]);

	pages.forEach((page, i) => {
		object(pageId(i), [
			ascii(
				`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${page.pointWidth} ${page.pointHeight}] ` +
					`/Resources << /XObject << /Im0 ${imageId(i)} 0 R >> >> ` +
					`/Contents ${contentId(i)} 0 R >>`,
			),
		]);

		// Stretch the image over the whole page: the transform is the page size,
		// so the raster's own resolution never has to be stated twice.
		const content = ascii(`q ${page.pointWidth} 0 0 ${page.pointHeight} 0 0 cm /Im0 Do Q\n`);
		object(contentId(i), [
			ascii(`<< /Length ${content.length} >>\nstream\n`),
			content,
			ascii("endstream"),
		]);

		const deflated = zlibSync(page.rgb, { level: 6 });
		object(imageId(i), [
			ascii(
				`<< /Type /XObject /Subtype /Image /Width ${page.width} /Height ${page.height} ` +
					`/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode ` +
					`/Length ${deflated.length} >>\nstream\n`,
			),
			deflated,
			ascii("\nendstream"),
		]);
	});

	const count = 3 + pages.length * 3;
	const startxref = length;
	push(ascii(`xref\n0 ${count}\n`));
	push(ascii("0000000000 65535 f \n"));
	for (let id = 1; id < count; id++) {
		push(ascii(`${String(offsets[id] ?? 0).padStart(10, "0")} 00000 n \n`));
	}
	push(ascii(`trailer\n<< /Size ${count} /Root 1 0 R >>\nstartxref\n${startxref}\n%%EOF\n`));

	return concat(chunks);
}

/** A filename-safe name for an exported deck, e.g. "deck.pdf". */
export function pdfFileName(deckName: string): string {
	const stem = deckName
		.replace(/\.pptx$/i, "")
		.replace(/[\\/:*?"<>|#^[\]]/g, "-")
		.trim();
	return `${stem || "deck"}.pdf`;
}
