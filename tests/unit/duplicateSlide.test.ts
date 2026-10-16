/**
 * Duplicating a slide, notes and all.
 *
 * A notes slide belongs to exactly one slide: its own relationships point back
 * at it. So the copy cannot simply share the original's notes — two slides
 * claiming one notes part is a file PowerPoint repairs on open — and dropping
 * the notes, which is what this used to do, loses the speaker's script for the
 * slide most likely to be presented next to its original. The copy gets a
 * notes part of its own, registered and related the way PowerPoint writes one.
 *
 * The other half is what the save writes: only what the duplicate needs,
 * leaving the original slide and its notes byte for byte as they were.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import { installDom } from "../../scripts/dom-shim";

installDom();

import { DeckEditor } from "../../src/edit/DeckEditor";
import { Selection } from "../../src/edit/Selection";
import { duplicateCurrentSlide } from "../../src/edit/slideCommands";
import { PptxPackage } from "../../src/pptx/package";
import { rebuildDeck } from "../../src/pptx/parse";
import { deleteSlide, duplicateSlide } from "../../src/ooxml/slideOps";
import { miniDeck, partText } from "./miniDeck";

const NOTES_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";

function overrideFor(pkg: PptxPackage, part: string): string | null {
	const types = partText(pkg, "[Content_Types].xml") ?? "";
	const match = new RegExp(`<Override[^>]*PartName="/${part.replace(/[.[\]]/g, "\\$&")}"[^>]*>`).exec(types);
	if (!match) return null;
	return /ContentType="([^"]+)"/.exec(match[0])?.[1] ?? null;
}

describe("duplicateSlide", () => {
	it("gives the copy notes of its own, with the same text", () => {
		const pkg = miniDeck([{ notes: "Say this first." }, {}]);
		const copy = duplicateSlide(pkg, "ppt/slides/slide1.xml", 1);

		const sourceNotes = pkg.relByKind("ppt/slides/slide1.xml", "notesSlide")?.target;
		const copyNotes = pkg.relByKind(copy, "notesSlide")?.target;
		assert.equal(sourceNotes, "ppt/notesSlides/notesSlide1.xml");
		assert.ok(copyNotes, "the copy has no notes");
		assert.notEqual(copyNotes, sourceNotes, "the copy shares the original's notes");
		assert.ok(pkg.has(copyNotes));

		const deck = rebuildDeck(pkg, "Deck");
		assert.equal(deck.slides[1].partPath, copy);
		assert.equal(deck.slides[1].notes, "Say this first.");
		assert.equal(deck.slides[0].notes, "Say this first.");
	});

	it("points the new notes back at the copy, and keeps the notes master", () => {
		const pkg = miniDeck([{ notes: "Script." }]);
		const copy = duplicateSlide(pkg, "ppt/slides/slide1.xml", 1);
		const copyNotes = pkg.relByKind(copy, "notesSlide")!.target;

		assert.equal(pkg.relByKind(copyNotes, "slide")?.target, copy);
		assert.equal(pkg.relByKind(copyNotes, "notesMaster")?.target, "ppt/notesMasters/notesMaster1.xml");
		// The original is still the original's.
		assert.equal(
			pkg.relByKind("ppt/notesSlides/notesSlide1.xml", "slide")?.target,
			"ppt/slides/slide1.xml",
		);
	});

	it("registers the new notes part's content type", () => {
		const pkg = miniDeck([{ notes: "Script." }]);
		const copy = duplicateSlide(pkg, "ppt/slides/slide1.xml", 1);
		const copyNotes = pkg.relByKind(copy, "notesSlide")!.target;
		assert.equal(overrideFor(pkg, copyNotes), NOTES_TYPE);
	});

	it("survives being saved and opened again", () => {
		const pkg = miniDeck([{ notes: "Script." }]);
		duplicateSlide(pkg, "ppt/slides/slide1.xml", 1);
		const reopened = PptxPackage.open(pkg.toZip());
		const deck = rebuildDeck(reopened, "Deck");
		assert.deepEqual(
			deck.slides.map((s) => s.notes),
			["Script.", "Script."],
		);
	});

	it("adds no notes to a copy of a slide that had none", () => {
		const pkg = miniDeck([{}]);
		const before = pkg.partPaths().filter((p) => p.startsWith("ppt/notesSlides/")).length;
		const copy = duplicateSlide(pkg, "ppt/slides/slide1.xml", 1);
		assert.equal(pkg.relByKind(copy, "notesSlide"), null);
		assert.equal(pkg.partPaths().filter((p) => p.startsWith("ppt/notesSlides/")).length, before);
	});

	it("takes the copy's notes, and only those, when the copy is deleted", () => {
		const pkg = miniDeck([{ notes: "Script." }]);
		const copy = duplicateSlide(pkg, "ppt/slides/slide1.xml", 1);
		const copyNotes = pkg.relByKind(copy, "notesSlide")!.target;
		assert.ok(deleteSlide(pkg, copy));
		assert.equal(pkg.has(copyNotes), false);
		assert.equal(pkg.has("ppt/notesSlides/notesSlide1.xml"), true);
		assert.equal(rebuildDeck(pkg, "Deck").slides[0].notes, "Script.");
	});
});

describe("duplicating from the editor", () => {
	function setup() {
		const pkg = miniDeck([{ notes: "Script." }, {}]);
		const deck = rebuildDeck(pkg, "Deck");
		const editor = new DeckEditor(pkg, { onChanged: () => undefined });
		const ctx = { editor, pkg, deck, slide: deck.slides[0], selection: new Selection() };
		return { pkg, ctx, editor };
	}

	it("undoes in one step, notes included", () => {
		const { pkg, ctx, editor } = setup();
		const before = new Map(pkg.partPaths().map((p) => [p, partText(pkg, p)]));
		assert.ok(duplicateCurrentSlide(ctx) >= 0);
		assert.ok(editor.undo());
		assert.deepEqual([...pkg.partPaths()].sort(), [...before.keys()].sort(), "parts left behind or lost");
		for (const [path, text] of before) assert.equal(partText(pkg, path), text, path);
	});

	it("leaves the original slide and its notes out of what the save rewrites", () => {
		const { pkg, ctx } = setup();
		pkg.clearDirty();
		duplicateCurrentSlide(ctx);
		const dirty = new Set(pkg.dirtyPartPaths);
		for (const untouched of [
			"ppt/slides/slide1.xml",
			"ppt/slides/_rels/slide1.xml.rels",
			"ppt/notesSlides/notesSlide1.xml",
			"ppt/notesSlides/_rels/notesSlide1.xml.rels",
		]) {
			assert.equal(dirty.has(untouched), false, `${untouched} would be rewritten`);
		}
	});
});
