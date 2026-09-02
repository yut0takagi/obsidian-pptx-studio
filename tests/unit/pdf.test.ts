/**
 * The PDF writer.
 *
 * A PDF that is almost right opens in nothing: the cross-reference table names
 * a byte offset for every object, and a reader that finds something other than
 * an object there gives up rather than guessing. That table cannot be checked
 * by eye, so it is checked here — every offset is followed to see what is
 * actually sitting at it.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import { unzlibSync } from "fflate";

import { buildPdf, pdfFileName, pointSize, rgbFromRgba, type PdfPage } from "../../src/export/pdf";

const decoder = new TextDecoder("latin1");

function page(width: number, height: number, fill = 0x40): PdfPage {
	return {
		width,
		height,
		pointWidth: width * 0.75,
		pointHeight: height * 0.75,
		rgb: new Uint8Array(width * height * 3).fill(fill),
	};
}

/**
 * The offsets the cross-reference table claims, in object order.
 *
 * The table is found by the newline in front of it, because "startxref" near
 * the end of the file contains the word too — and matching that one instead
 * finds no entries at all, which reads as a passing test rather than a broken
 * search.
 */
function xrefOffsets(pdf: Uint8Array): number[] {
	const text = decoder.decode(pdf);
	const start = text.lastIndexOf("\nxref\n");
	assert.notEqual(start, -1, "no cross-reference table in the file");
	const body = text.slice(start);
	const entries = Array.from(body.matchAll(/^(\d{10}) \d{5} [nf] $/gm)).map((m) => Number(m[1]));
	assert.ok(entries.length > 0, "the cross-reference table has no entries");
	return entries;
}

describe("buildPdf", () => {
	it("writes something a reader will recognise as a PDF", () => {
		const text = decoder.decode(buildPdf([page(4, 4)]));
		assert.ok(text.startsWith("%PDF-1.4\n"), "header");
		assert.ok(text.trimEnd().endsWith("%%EOF"), "trailer");
	});

	it("points every cross-reference entry at the object it names", () => {
		const pdf = buildPdf([page(4, 4), page(4, 4)]);
		const text = decoder.decode(pdf);
		const offsets = xrefOffsets(pdf);
		// Entry 0 is the head of the free list and points nowhere.
		for (let id = 1; id < offsets.length; id++) {
			assert.ok(
				text.startsWith(`${id} 0 obj`, offsets[id]),
				`object ${id} is not at the offset the xref gives (${offsets[id]})`,
			);
		}
	});

	it("agrees with itself about how many objects there are", () => {
		const pdf = buildPdf([page(4, 4), page(4, 4), page(4, 4)]);
		const text = decoder.decode(pdf);
		const size = Number(/\/Size (\d+)/.exec(text)?.[1]);
		assert.equal(size, 3 + 3 * 3);
		assert.equal(xrefOffsets(pdf).length, size);
		assert.equal(text.match(/\/Type \/Page\b/g)?.length, 3);
	});

	it("puts startxref where the table actually begins", () => {
		const pdf = buildPdf([page(4, 4)]);
		const text = decoder.decode(pdf);
		const at = Number(/startxref\n(\d+)/.exec(text)?.[1]);
		assert.ok(text.startsWith("xref\n", at), "startxref does not land on the table");
	});

	it("lists each page in the page tree", () => {
		const text = decoder.decode(buildPdf([page(4, 4), page(4, 4)]));
		assert.match(text, /\/Type \/Pages \/Count 2 \/Kids \[3 0 R 6 0 R\]/);
	});

	it("gives the page the size it was asked for", () => {
		const p = page(1920, 1080);
		const text = decoder.decode(buildPdf([p]));
		assert.ok(text.includes(`/MediaBox [0 0 ${p.pointWidth} ${p.pointHeight}]`));
	});

	it("stores the pixels losslessly", () => {
		const p = page(3, 2);
		p.rgb = new Uint8Array([
			1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18,
		]);
		const pdf = buildPdf([p]);
		const text = decoder.decode(pdf);
		// The image stream is the last one in the file, after its own dictionary.
		const dict = text.lastIndexOf("/Subtype /Image");
		const from = text.indexOf("stream\n", dict) + "stream\n".length;
		const declared = Number(/\/Length (\d+) >>\nstream\n$/.exec(text.slice(dict, from))?.[1]);
		assert.ok(declared > 0, "the image stream declares no length");
		assert.deepEqual(Array.from(unzlibSync(pdf.slice(from, from + declared))), Array.from(p.rgb));
	});

	it("declares a stream length that matches the bytes written", () => {
		const text = decoder.decode(buildPdf([page(4, 4)]));
		const content = /\/Length (\d+) >>\nstream\n(.*?)endstream/s.exec(text);
		assert.ok(content);
		assert.equal(Number(content[1]), content[2].length);
	});

	it("still produces a well-formed file with no pages at all", () => {
		const text = decoder.decode(buildPdf([]));
		assert.match(text, /\/Type \/Pages \/Count 0 \/Kids \[\]/);
		assert.match(text, /\/Size 3/);
		assert.ok(text.trimEnd().endsWith("%%EOF"));
	});
});

describe("pointSize", () => {
	it("converts CSS pixels to points", () => {
		assert.deepEqual(pointSize(960, 540), { width: 720, height: 405 });
	});

	it("keeps a widescreen slide's proportions", () => {
		const { width, height } = pointSize(1280, 720);
		assert.ok(Math.abs(width / height - 16 / 9) < 1e-9);
	});
});

describe("rgbFromRgba", () => {
	it("drops the alpha channel and keeps the order", () => {
		const rgba = new Uint8Array([1, 2, 3, 255, 4, 5, 6, 128]);
		assert.deepEqual(Array.from(rgbFromRgba(rgba)), [1, 2, 3, 4, 5, 6]);
	});

	it("returns nothing for nothing", () => {
		assert.equal(rgbFromRgba(new Uint8Array([])).length, 0);
	});
});

describe("pdfFileName", () => {
	it("swaps the extension", () => {
		assert.equal(pdfFileName("Quarterly review.pptx"), "Quarterly review.pdf");
	});

	it("takes out what a filesystem would object to", () => {
		assert.equal(pdfFileName("a/b:c*d.pptx"), "a-b-c-d.pdf");
	});

	it("falls back rather than producing a nameless file", () => {
		assert.equal(pdfFileName(""), "deck.pdf");
	});
});
