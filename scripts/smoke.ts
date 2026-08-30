/**
 * Node smoke test for the parser and the save round trip.
 *
 * The renderer needs a browser, but everything up to it — unzip, relationships,
 * theme resolution, placeholder inheritance, text extraction and repackaging —
 * is pure data manipulation, so it can be exercised against real decks here.
 *
 *   npm run smoke -- ~/Downloads/*.pptx
 */
import { readFileSync, writeFileSync } from "node:fs";
import { basename } from "node:path";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

installDomShims();

import { parseDeck } from "../src/pptx/parse";
import type { Shape, TextBody } from "../src/pptx/types";

/** xmldom implements the core DOM but not the element-traversal conveniences. */
function installDomShims(): void {
	const g = globalThis as Record<string, unknown>;
	g.DOMParser = DOMParser;
	g.XMLSerializer = XMLSerializer;

	const probe = new DOMParser().parseFromString("<a/>", "application/xml");
	const proto = Object.getPrototypeOf(probe.documentElement) as object;
	if (!("firstElementChild" in proto)) {
		Object.defineProperties(proto, {
			firstElementChild: {
				get(this: Node) {
					for (let n = this.firstChild; n; n = n.nextSibling) {
						if (n.nodeType === 1) return n;
					}
					return null;
				},
			},
			nextElementSibling: {
				get(this: Node) {
					for (let n = this.nextSibling; n; n = n.nextSibling) {
						if (n.nodeType === 1) return n;
					}
					return null;
				},
			},
		});
	}

	if (typeof (g.URL as { createObjectURL?: unknown }).createObjectURL !== "function") {
		(g.URL as { createObjectURL: (b: unknown) => string }).createObjectURL = () =>
			"blob:stub";
		(g.URL as unknown as { revokeObjectURL: () => void }).revokeObjectURL = () => undefined;
	}
}

function countShapes(shapes: Shape[]): number {
	let total = 0;
	for (const shape of shapes) {
		total += 1;
		if (shape.kind === "group") total += countShapes(shape.children);
	}
	return total;
}

function collectText(shapes: Shape[]): string[] {
	const out: string[] = [];
	const push = (body: TextBody | null) => {
		if (!body) return;
		for (const p of body.paragraphs) {
			const text = p.runs.map((r) => r.text).join("").trim();
			if (text) out.push(text);
		}
	};
	for (const shape of shapes) {
		if (shape.kind === "shape") push(shape.text);
		else if (shape.kind === "table") {
			for (const row of shape.table.rows) for (const cell of row.cells) push(cell.text);
		} else if (shape.kind === "group") out.push(...collectText(shape.children));
	}
	return out;
}

function check(path: string): boolean {
	const name = basename(path);
	let ok = true;
	try {
		const { deck, pkg } = parseDeck(readFileSync(path), name);
		const shapes = deck.slides.reduce((n, s) => n + countShapes(s.shapes), 0);
		const text = deck.slides.flatMap((s) => collectText(s.shapes.slice(s.templateShapes)));
		const chars = text.join("").length;
		const notes = deck.slides.filter((s) => s.notes).length;

		console.log(
			`  ${name}\n` +
				`    ${deck.slides.length} slides · ${Math.round(deck.width)}x${Math.round(deck.height)}px · ` +
				`${shapes} shapes · ${chars} chars of text · ${notes} with notes`,
		);

		if (deck.slides.length === 0) {
			console.log("    FAIL: no slides parsed");
			ok = false;
		}
		if (chars === 0) {
			console.log("    WARN: no text extracted");
		}
		const sample = text.find((t) => t.length > 4);
		if (sample) console.log(`    first text: ${JSON.stringify(sample.slice(0, 60))}`);

		// Round trip: repackage untouched and confirm the result still parses.
		const rezipped = pkg.toZip();
		const reopened = parseDeck(rezipped, name);
		if (reopened.deck.slides.length !== deck.slides.length) {
			console.log(
				`    FAIL: round trip changed slide count ` +
					`(${deck.slides.length} -> ${reopened.deck.slides.length})`,
			);
			ok = false;
		}
		reopened.pkg.dispose();
		pkg.dispose();
	} catch (error) {
		console.log(`  ${name}\n    FAIL: ${(error as Error).message}`);
		ok = false;
	}
	return ok;
}

/** Edit the first run of the first text box, save, reopen and verify. */
function checkEditRoundTrip(path: string): boolean {
	const name = basename(path);
	const { deck, pkg } = parseDeck(readFileSync(path), name);
	try {
		for (const slide of deck.slides) {
			for (const shape of slide.shapes.slice(slide.templateShapes)) {
				if (shape.kind !== "shape" || !shape.text?.source) continue;
				const para = shape.text.paragraphs.find((p) => p.runs.some((r) => r.source));
				const run = para?.runs.find((r) => r.source);
				if (!run?.source) continue;

				const marker = "SMOKE-TEST-EDIT";
				const target = run.source.getElementsByTagName("a:t")[0];
				if (!target) continue;
				target.textContent = marker;
				pkg.markDirty(shape.text.sourcePart);

				const out = join(tmpdir(), `pptx-smoke-${Date.now()}.pptx`);
				writeFileSync(out, pkg.toZip());
				const reopened = parseDeck(readFileSync(out), name);
				const found = reopened.deck.slides.some((s) =>
					collectText(s.shapes).some((t) => t.includes(marker)),
				);
				reopened.pkg.dispose();
				console.log(
					found
						? `    edit round trip: OK (wrote ${marker} into ${shape.text.sourcePart})`
						: `    FAIL: edit did not survive the round trip`,
				);
				return found;
			}
		}
		console.log("    edit round trip: skipped (no editable text found)");
		return true;
	} finally {
		pkg.dispose();
	}
}

const files = process.argv.slice(2);
if (files.length === 0) {
	console.error("usage: npm run smoke -- <file.pptx> [...]");
	process.exit(2);
}

console.log(`Checking ${files.length} deck(s)\n`);
let failures = 0;
for (const file of files) {
	if (!check(file)) failures++;
}
console.log("\nEdit round trip on the first deck:");
if (!checkEditRoundTrip(files[0])) failures++;

console.log(failures === 0 ? "\nAll checks passed." : `\n${failures} check(s) failed.`);
process.exit(failures === 0 ? 0 : 1);
