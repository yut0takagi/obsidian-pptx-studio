/**
 * Bullets on paragraphs that have nothing to say.
 *
 * Decks space their content out with empty paragraphs, and those paragraphs
 * inherit the master's `a:buChar` like any other. Parsing that inheritance
 * literally puts a bullet glyph on a blank line — a dot sitting in the
 * whitespace, next to nothing. PowerPoint draws no such thing, so neither
 * should we.
 *
 * The deck that surfaced this had 84 of them across 34 slides, all from a
 * single `<a:buChar char="•"/>` on lvl1 of the master.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import { installDomParser } from "./dom";

installDomParser();

import { parseTextBody } from "../../src/pptx/text";
import type { ParseContext } from "../../src/pptx/style";
import type { TextBody } from "../../src/pptx/types";

const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

/** A master-style list that puts a bullet on every level-1 paragraph. */
const MASTER_LIST = `<a:lstStyle ${A}><a:lvl1pPr><a:buChar char="•"/></a:lvl1pPr></a:lstStyle>`;

function el(xml: string): Element {
	const doc = new DOMParser().parseFromString(xml, "application/xml");
	const node = doc.documentElement;
	assert.ok(node && node.nodeName !== "parsererror", `could not parse: ${xml}`);
	return node;
}

/**
 * The parser reaches into the package only for theme colours and relationship
 * lookups, neither of which any case here needs, so a bare context is enough.
 */
function ctx(): ParseContext {
	return {
		pkg: null as never,
		partPath: "ppt/slides/slide1.xml",
		theme: { colors: {}, fontMajor: null, fontMinor: null } as never,
	};
}

function parse(body: string, chain: string[] = [MASTER_LIST]): TextBody {
	const result = parseTextBody(el(body), ctx(), chain.map(el), null);
	assert.ok(result, "expected a text body");
	return result;
}

function body(...paragraphs: string[]): string {
	return `<a:txBody ${A}><a:bodyPr/>${paragraphs.join("")}</a:txBody>`;
}

describe("bullets on empty paragraphs", () => {
	it("leaves a paragraph with no runs unbulleted", () => {
		// The spacer PowerPoint writes: an end-properties element and nothing else.
		const text = parse(body("<a:p><a:endParaRPr/></a:p>"));

		assert.equal(text.paragraphs.length, 1);
		assert.equal(text.paragraphs[0].bullet, null);
	});

	it("leaves a paragraph whose runs are all empty unbulleted", () => {
		// Same blank line on screen, written a different way — an editor that
		// leaves the run behind after the text is deleted.
		const text = parse(body("<a:p><a:r><a:t></a:t></a:r></a:p>"));

		assert.equal(text.paragraphs[0].bullet, null);
	});

	it("still bullets a paragraph that has text", () => {
		const text = parse(body("<a:p><a:r><a:t>Point</a:t></a:r></a:p>"));

		assert.equal(text.paragraphs[0].bullet?.text, "•");
	});

	it("bullets only the paragraphs with text when they are interleaved", () => {
		const text = parse(
			body(
				"<a:p><a:r><a:t>First</a:t></a:r></a:p>",
				"<a:p><a:endParaRPr/></a:p>",
				"<a:p><a:r><a:t>Second</a:t></a:r></a:p>",
			),
		);

		assert.deepEqual(
			text.paragraphs.map((p) => p.bullet?.text ?? null),
			["•", null, "•"],
		);
	});

	it("does not let a blank line consume an ordinal", () => {
		// A number skipped here would show up as "1. 3." in the rendered list,
		// which is worse than the stray dot this all started with.
		const numbered = `<a:lstStyle ${A}><a:lvl1pPr><a:buAutoNum type="arabicPeriod"/></a:lvl1pPr></a:lstStyle>`;
		const text = parse(
			body(
				"<a:p><a:r><a:t>First</a:t></a:r></a:p>",
				"<a:p><a:endParaRPr/></a:p>",
				"<a:p><a:r><a:t>Second</a:t></a:r></a:p>",
			),
			[numbered],
		);

		assert.deepEqual(
			text.paragraphs.map((p) => p.bullet?.text ?? null),
			["1.", null, "2."],
		);
	});

	it("keeps the empty paragraph, so the spacing it exists for survives", () => {
		const text = parse(body("<a:p><a:r><a:t>Point</a:t></a:r></a:p>", "<a:p><a:endParaRPr/></a:p>"));

		assert.equal(text.paragraphs.length, 2);
		assert.deepEqual(text.paragraphs[1].runs.map((r) => r.text), [""]);
	});
});
