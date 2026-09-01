/**
 * The XML helpers every other parser sits on.
 *
 * The point of these is prefix independence: a deck written by Keynote or
 * Google Slides binds the same namespaces to different prefixes, so matching on
 * `a:off` rather than the local name `off` would parse PowerPoint's output and
 * nothing else. Each case here is written with a prefix that is deliberately
 * not the one PowerPoint emits.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import { installDomParser } from "./dom";

installDomParser();

import {
	attr,
	boolAttr,
	child,
	childPath,
	children,
	descendant,
	descendants,
	numAttr,
	parseXml,
} from "../../src/pptx/xml";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

/** Parse a fragment and hand back its root element. */
function el(xml: string): Element {
	const doc = new DOMParser().parseFromString(xml, "application/xml");
	return doc.documentElement;
}

function bytes(text: string): Uint8Array {
	return new TextEncoder().encode(text);
}

describe("parseXml", () => {
	it("parses a document", () => {
		const doc = parseXml(bytes(`<?xml version="1.0"?><root><a/></root>`));
		assert.equal(doc.documentElement.localName, "root");
	});

	it("tolerates a byte order mark", () => {
		const raw = bytes(`<?xml version="1.0"?><root/>`);
		const withBom = new Uint8Array(raw.length + 3);
		withBom.set([0xef, 0xbb, 0xbf]);
		withBom.set(raw, 3);
		assert.equal(parseXml(withBom).documentElement.localName, "root");
	});

	it("throws on malformed XML rather than returning a parsererror document", () => {
		assert.throws(() => parseXml(bytes("<root><unclosed></root>")), /Malformed XML/);
	});
});

describe("child / children", () => {
	const sp = el(`<x:sp xmlns:x="${A}">
		<x:spPr><x:xfrm/></x:spPr>
		<x:txBody><x:p/><x:p/></x:txBody>
	</x:sp>`);

	it("matches on the local name, whatever the prefix", () => {
		assert.equal(child(sp, "spPr")?.localName, "spPr");
	});

	it("only looks at direct children", () => {
		assert.equal(child(sp, "xfrm"), null);
		assert.equal(children(sp, "p").length, 0);
	});

	it("returns every direct match in document order", () => {
		const body = child(sp, "txBody");
		assert.equal(children(body, "p").length, 2);
	});

	it("treats a missing element as empty rather than throwing", () => {
		assert.equal(child(null, "spPr"), null);
		assert.deepEqual(children(undefined, "p"), []);
	});
});

describe("childPath", () => {
	const sp = el(`<x:sp xmlns:x="${A}"><x:spPr><x:xfrm><x:off x="1"/></x:xfrm></x:spPr></x:sp>`);

	it("walks a chain of direct children", () => {
		assert.equal(attr(childPath(sp, "spPr", "xfrm", "off"), "x"), "1");
	});

	it("stops at the first missing link", () => {
		assert.equal(childPath(sp, "spPr", "nope", "off"), null);
	});
});

describe("descendant / descendants", () => {
	const tree = el(`<x:root xmlns:x="${A}">
		<x:a><x:t>first</x:t></x:a>
		<x:b><x:c><x:t>second</x:t></x:c></x:b>
	</x:root>`);

	it("finds matches at any depth, in document order", () => {
		assert.deepEqual(
			descendants(tree, "t").map((n) => n.textContent),
			["first", "second"],
		);
	});

	it("returns the first match for the singular form", () => {
		assert.equal(descendant(tree, "t")?.textContent, "first");
	});

	it("returns null when nothing matches", () => {
		assert.equal(descendant(tree, "missing"), null);
	});
});

describe("attr", () => {
	const node = el(`<x:blip xmlns:x="${A}" xmlns:r="http://x/rel" r:embed="rId3" cx="42"/>`);

	it("reads an unprefixed attribute", () => {
		assert.equal(attr(node, "cx"), "42");
	});

	it("reads a prefixed attribute by its local name", () => {
		assert.equal(attr(node, "embed"), "rId3");
	});

	it("returns null when the attribute is absent", () => {
		assert.equal(attr(node, "nope"), null);
		assert.equal(attr(null, "cx"), null);
	});
});

describe("numAttr", () => {
	const node = el(`<x:ext xmlns:x="${A}" cx="9525" neg="-1" frac="1.5" empty="" bad="abc"/>`);

	it("parses numbers, including negative and fractional ones", () => {
		assert.equal(numAttr(node, "cx"), 9525);
		assert.equal(numAttr(node, "neg"), -1);
		assert.equal(numAttr(node, "frac"), 1.5);
	});

	it("returns null rather than NaN or 0 for empty and unparseable values", () => {
		assert.equal(numAttr(node, "empty"), null);
		assert.equal(numAttr(node, "bad"), null);
		assert.equal(numAttr(node, "missing"), null);
	});
});

describe("boolAttr", () => {
	const node = el(`<x:xfrm xmlns:x="${A}" a="1" b="true" c="0" d="false" e=""/>`);

	it('accepts both "1" and "true"', () => {
		assert.equal(boolAttr(node, "a"), true);
		assert.equal(boolAttr(node, "b"), true);
	});

	it('accepts both "0" and "false"', () => {
		assert.equal(boolAttr(node, "c"), false);
		assert.equal(boolAttr(node, "d"), false);
	});

	it("reads a present-but-empty attribute as true", () => {
		assert.equal(boolAttr(node, "e"), true);
	});

	it("distinguishes absent from false", () => {
		assert.equal(boolAttr(node, "missing"), null);
	});
});
