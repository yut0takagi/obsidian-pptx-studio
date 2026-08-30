/**
 * Smoke test for the parser, the renderer and the edit write-back.
 *
 *   npm run smoke -- ~/Downloads/*.pptx
 *
 * The edit test goes through the same path the UI does — render the slide,
 * change text in the rendered DOM, commit — rather than editing XML directly,
 * so it would catch a break anywhere between the two.
 */
import { readFileSync, writeFileSync } from "node:fs";
import { basename, join } from "node:path";
import { tmpdir } from "node:os";
import { installDom } from "./dom-shim";

installDom();

import { commitTextBody } from "../src/edit/textEdit";
import { writeCrop } from "../src/edit/CropController";
import { guideParts, readGuides, writeGuides } from "../src/ooxml/guides";
import { DeckEditor } from "../src/edit/DeckEditor";
import { Selection } from "../src/edit/Selection";
import {
	type CommandContext,
	alignSelection,
	deleteSelection,
	duplicateSelection,
	groupSelection,
	renameShape,
	reorderSelection,
	setShapeHidden,
	ungroupSelection,
} from "../src/edit/commands";
import {
	applyTextFormat,
	changeShape,
	copyFormatting,
	flipSelection,
	pasteFormatting,
	rotateBy,
	setFill,
	setGeometry,
	setHyperlink,
} from "../src/edit/formatCommands";
import {
	TableSelection,
	deleteTableColumns,
	deleteTableRows,
	insertTableColumn,
	insertTableRow,
	mergeTableCells,
	setCellFill,
	splitTableCells,
} from "../src/edit/tableCommands";
import { insertAutoShape, insertPicture, insertTable, insertTextBox } from "../src/edit/insertCommands";
import {
	deleteCurrentSlide,
	duplicateCurrentSlide,
	newSlide,
	reorderSlide,
	setSlideBackgroundColor,
} from "../src/edit/slideCommands";
import { encodePng } from "./png.mjs";
import { writeFrame, writeShapeFrame } from "../src/edit/geometryWrite";
import { parseDeck, rebuildDeck } from "../src/pptx/parse";
import type { PptxPackage } from "../src/pptx/package";
import { renderSlide } from "../src/render/renderSlide";
import { shapeRegistry, textBodyRegistry } from "../src/render/renderSlide";
import type { Deck, Shape, Slide, TextBody } from "../src/pptx/types";

let failures = 0;

function fail(message: string): void {
	console.log(`    FAIL: ${message}`);
	failures++;
}

function pass(message: string): void {
	console.log(`    ok: ${message}`);
}

// ------------------------------------------------------------------ helpers

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

function deckText(deck: Deck): string {
	return deck.slides.flatMap((s) => collectText(s.shapes)).join(" ");
}

/**
 * Render every slide and return the first editable text box matching a
 * predicate, so each test can ask for the shape of content it needs.
 */
function findEditableBox(
	deck: Deck,
	predicate: (body: TextBody, box: HTMLElement) => boolean,
): { slide: Slide; box: HTMLElement } | null {
	for (const slide of deck.slides) {
		const el = renderSlide(deck, slide);
		for (const box of Array.from(el.querySelectorAll<HTMLElement>('[data-editable="1"]'))) {
			const body = textBodyRegistry.get(box);
			if (body && predicate(body, box)) return { slide, box };
		}
	}
	return null;
}

const hasText = (body: TextBody) =>
	body.paragraphs.some((p) => p.runs.some((r) => r.text.trim() !== ""));

function reopen(pkg: PptxPackage, name: string): Deck {
	const bytes = pkg.toZip();
	const out = join(tmpdir(), `pptx-smoke-${Date.now()}-${Math.random().toString(36).slice(2)}.pptx`);
	writeFileSync(out, bytes);
	const { deck, pkg: reopened } = parseDeck(readFileSync(out), name);
	reopened.dispose();
	return deck;
}

// -------------------------------------------------------------------- tests

function checkDeck(path: string): void {
	const name = basename(path);
	console.log(`  ${name}`);
	try {
		const { deck, pkg } = parseDeck(readFileSync(path), name);
		const shapes = deck.slides.reduce((n, s) => n + countShapes(s.shapes), 0);
		const text = deck.slides.flatMap((s) => collectText(s.shapes.slice(s.templateShapes)));
		const chars = text.join("").length;

		console.log(
			`    ${deck.slides.length} slides · ${Math.round(deck.width)}x${Math.round(deck.height)}px · ` +
				`${shapes} shapes · ${chars} chars · ${deck.slides.filter((s) => s.notes).length} with notes`,
		);

		if (deck.slides.length === 0) fail("no slides parsed");

		// Rendering must not throw, and must produce something for every slide.
		for (const slide of deck.slides) {
			const el = renderSlide(deck, slide);
			if (slide.shapes.length > 0 && el.childElementCount === 0) {
				fail(`slide ${slide.index} rendered no elements from ${slide.shapes.length} shapes`);
			}
		}

		const rebuilt = reopen(pkg, name);
		if (rebuilt.slides.length !== deck.slides.length) {
			fail(`repackaging changed the slide count (${deck.slides.length} -> ${rebuilt.slides.length})`);
		}
		pkg.dispose();
	} catch (error) {
		fail((error as Error).message);
	}
}

/** Edit one run through the rendered DOM and confirm it survives a save. */
function checkRunEdit(path: string): void {
	const name = basename(path);
	const { deck, pkg } = parseDeck(readFileSync(path), name);
	try {
		// Prefer a paragraph with several runs: that is where run matching matters.
		const found =
			findEditableBox(
				deck,
				(body) =>
					hasText(body) && body.paragraphs.some((p) => p.runs.filter((r) => r.text.trim()).length > 1),
			) ?? findEditableBox(deck, hasText);
		if (!found) {
			console.log("    skipped: no editable text");
			return;
		}
		const box = found.box;
		const body = textBodyRegistry.get(box);
		const spans = Array.from(box.querySelectorAll<HTMLElement>("[data-run]")).filter(
			(s) => (s.textContent ?? "").trim() !== "",
		);
		// Edit inside a paragraph that has neighbours, so the preservation check bites.
		const span =
			spans.find(
				(s) =>
					Array.from(
						(s.parentElement as HTMLElement).querySelectorAll<HTMLElement>("[data-run]"),
					).filter((n) => (n.textContent ?? "").trim() !== "").length > 1,
			) ?? spans[0];
		if (!body || !span) {
			console.log("    skipped: no run to edit");
			return;
		}

		const siblings = Array.from(
			(span.parentElement as HTMLElement).querySelectorAll<HTMLElement>("[data-run]"),
		)
			.filter((s) => s !== span)
			.map((s) => s.textContent ?? "");

		const marker = "EDITED-RUN";
		span.textContent = marker;

		const result = commitTextBody(box);
		if (!result.changed || !result.part) {
			fail("commitTextBody reported no change after editing a run");
			return;
		}
		pkg.markDirty(result.part);

		const saved = reopen(pkg, name);
		if (!deckText(saved).includes(marker)) fail("the edited run did not survive the save");
		else pass(`edited run written into ${result.part} and read back`);

		// PPTX_SMOKE_OUT lets a caller keep the edited deck so it can be opened in
		// a real OOXML consumer, which is the only check that matters in the end.
		const keep = process.env.PPTX_SMOKE_OUT;
		if (keep) {
			writeFileSync(keep, pkg.toZip());
			console.log(`    wrote edited deck to ${keep}`);
		}

		// Runs the user did not touch must come back unchanged: that is what keeps
		// mixed formatting inside a paragraph intact.
		const savedText = deckText(saved);
		const lost = siblings.filter((t) => t.trim() !== "" && !savedText.includes(t.trim()));
		if (lost.length > 0) fail(`sibling runs were lost: ${JSON.stringify(lost)}`);
		else if (siblings.length > 0) pass(`${siblings.length} untouched sibling run(s) preserved`);
	} finally {
		pkg.dispose();
	}
}

/** Delete a paragraph in the DOM and confirm the a:p is removed from the XML. */
function checkParagraphDelete(path: string): void {
	const name = basename(path);
	const { deck, pkg } = parseDeck(readFileSync(path), name);
	try {
		const found = findEditableBox(
			deck,
			(body) => hasText(body) && body.paragraphs.filter((p) => p.runs.some((r) => r.text.trim())).length > 1,
		);
		if (!found) {
			console.log("    skipped: no multi-paragraph text box");
			return;
		}
		const box = found.box;
		const before = textBodyRegistry.get(box)!.paragraphs.length;
		const removed = box.children[before - 1];
		if (!removed) {
			console.log("    skipped: paragraph elements and model are out of step");
			return;
		}
		const removedText = (removed.textContent ?? "").trim();
		removed.remove();

		const result = commitTextBody(box);
		if (!result.part) {
			fail("deleting a paragraph reported no change");
			return;
		}
		pkg.markDirty(result.part);

		const saved = reopen(pkg, name);
		const savedBox = saved.slides
			.flatMap((s) => s.shapes)
			.find((s) => s.kind === "shape" && s.text?.sourcePart === result.part && s.text.paragraphs.length === before - 1);

		if (!savedBox) fail(`paragraph count did not drop from ${before} to ${before - 1}`);
		else if (removedText && deckText(saved).includes(removedText)) {
			fail("the deleted paragraph's text is still present");
		} else pass(`paragraph removed (${before} -> ${before - 1})`);
	} finally {
		pkg.dispose();
	}
}

/**
 * Charts must come back with their cached numbers intact and render to real SVG.
 *
 * macOS Quick Look draws nothing at all for a chart, so there is no external
 * renderer to compare against here — the values and the emitted geometry are the
 * check.
 */
function checkCharts(path: string): void {
	const name = basename(path);
	const { deck, pkg } = parseDeck(readFileSync(path), name);
	try {
		const charts = deck.slides.flatMap((slide) =>
			slide.shapes.filter((shape): shape is Extract<Shape, { kind: "chart" }> => shape.kind === "chart"),
		);
		if (charts.length === 0) {
			console.log("    skipped: no charts in this deck");
			return;
		}

		for (const shape of charts) {
			const chart = shape.chart;
			const points = chart.series.reduce(
				(total, series) => total + series.values.filter((v) => v !== null).length,
				0,
			);
			console.log(
				`    ${chart.kind}: ${chart.series.length} series x ${chart.categories.length} categories ` +
					`(${points} points)${chart.title ? ` "${chart.title}"` : ""}`,
			);
			if (chart.empty) fail(`chart "${chart.title}" parsed as empty`);
			else if (points === 0) fail(`chart "${chart.title}" has no values`);
		}

		// The fixture's own chart, checked value by value.
		const column = charts.find((c) => c.chart.kind === "column");
		if (column) {
			const expected = [
				{ name: "2024", values: [32, 41, 38, 55] },
				{ name: "2025", values: [45, 39, 52, 68] },
			];
			const actual = column.chart.series.map((s) => ({ name: s.name, values: s.values }));
			if (JSON.stringify(actual) === JSON.stringify(expected)) {
				pass("column chart values round-tripped exactly");
			} else {
				fail(`column chart values came back as ${JSON.stringify(actual)}`);
			}
			if (column.chart.categories.join(",") !== "Q1,Q2,Q3,Q4") {
				fail(`column chart categories came back as ${column.chart.categories.join(",")}`);
			} else pass("column chart categories round-tripped");
		}

		// Rendering must emit real geometry, not just a label.
		const slide = deck.slides.find((s) => s.shapes.some((sh) => sh.kind === "chart"));
		if (slide) {
			const el = renderSlide(deck, slide);
			const bars = el.querySelectorAll("rect").length;
			const wedges = el.querySelectorAll("path").length;
			const labels = el.querySelectorAll("text").length;
			if (bars > 0 && wedges > 0 && labels > 0) {
				pass(`rendered ${bars} bars, ${wedges} wedges and ${labels} labels`);
			} else {
				fail(`chart rendering produced ${bars} bars, ${wedges} wedges, ${labels} labels`);
			}
		}
	} finally {
		pkg.dispose();
	}
}

/** Move and resize shapes through the same call the drag handler makes. */
function checkGeometryEdit(path: string): void {
	const name = basename(path);
	const { deck, pkg } = parseDeck(readFileSync(path), name);
	try {
		// A shape with its own a:xfrm, and a placeholder that inherits one: the
		// second is the interesting case, since moving it must create the element.
		let explicit: { shape: Shape; slide: Slide } | null = null;
		let inherited: { shape: Shape; slide: Slide } | null = null;

		for (const slide of deck.slides) {
			const el = renderSlide(deck, slide);
			for (const node of Array.from(el.querySelectorAll<HTMLElement>('[data-selectable="1"]'))) {
				const shape = shapeRegistry.get(node);
				if (!shape?.source) continue;
				const hasXfrm = shape.source.getElementsByTagName("a:xfrm").length > 0;
				if (hasXfrm && !explicit) explicit = { shape, slide };
				if (!hasXfrm && !inherited) inherited = { shape, slide };
			}
			if (explicit && inherited) break;
		}

		for (const [label, found] of [
			["explicit a:xfrm", explicit],
			["inherited frame", inherited],
		] as const) {
			if (!found) {
				console.log(`    skipped: no shape with an ${label}`);
				continue;
			}
			const { shape, slide } = found;
			const target = { ...shape.frame, x: 111, y: 222, w: 333, h: 144 };
			const part = writeShapeFrame(shape, target);
			if (!part) {
				fail(`writeShapeFrame refused a ${shape.kind} with an ${label}`);
				continue;
			}
			pkg.markDirty(part);

			const saved = reopen(pkg, name);
			const savedShape = saved.slides[slide.index - 1]?.shapes.find((s) => s.id === shape.id);
			if (!savedShape) {
				fail(`shape ${shape.id} disappeared after moving it (${label})`);
				continue;
			}
			const f = savedShape.frame;
			const close = (a: number, b: number) => Math.abs(a - b) < 0.5;
			if (close(f.x, 111) && close(f.y, 222) && close(f.w, 333) && close(f.h, 144)) {
				pass(`${label}: ${shape.kind} moved and resized, read back exactly`);
			} else {
				fail(
					`${label}: expected 111,222 333x144 but read back ` +
						`${f.x.toFixed(1)},${f.y.toFixed(1)} ${f.w.toFixed(1)}x${f.h.toFixed(1)}`,
				);
			}
		}
	} finally {
		pkg.dispose();
	}
}

/** Undo must put the XML back exactly, and redo must put the edit back. */
function checkUndoRedo(path: string): void {
	const name = basename(path);
	const { deck, pkg } = parseDeck(readFileSync(path), name);
	try {
		let subject: { shape: Shape; slide: Slide } | null = null;
		for (const slide of deck.slides) {
			const el = renderSlide(deck, slide);
			for (const node of Array.from(el.querySelectorAll<HTMLElement>('[data-selectable="1"]'))) {
				const shape = shapeRegistry.get(node);
				if (shape?.source && shape.frame.w > 0) {
					subject = { shape, slide };
					break;
				}
			}
			if (subject) break;
		}
		if (!subject) {
			console.log("    skipped: no selectable shape");
			return;
		}

		const { shape, slide } = subject;
		const original = { ...shape.frame };
		const editor = new DeckEditor(pkg, { onChanged: () => undefined });
		const before = editor.capture([shape.sourcePart]);

		const part = writeShapeFrame(shape, { ...shape.frame, x: 500, y: 400 });
		if (!part) {
			fail("writeShapeFrame refused the subject shape");
			return;
		}
		editor.recordApplied("Move shape", before);

		const read = (): Shape | undefined =>
			reopen(pkg, name).slides[slide.index - 1]?.shapes.find((s) => s.id === shape.id);

		const moved = read();
		if (!moved || Math.abs(moved.frame.x - 500) > 0.5) {
			fail("the move did not take effect");
			return;
		}

		editor.undo();
		const undone = read();
		if (!undone) {
			fail("the shape disappeared after undo");
			return;
		}
		if (Math.abs(undone.frame.x - original.x) > 0.5 || Math.abs(undone.frame.y - original.y) > 0.5) {
			fail(
				`undo left the shape at ${undone.frame.x.toFixed(1)},${undone.frame.y.toFixed(1)} ` +
					`instead of ${original.x.toFixed(1)},${original.y.toFixed(1)}`,
			);
			return;
		}
		pass("undo restored the original position");

		editor.redo();
		const redone = read();
		if (!redone || Math.abs(redone.frame.x - 500) > 0.5) fail("redo did not reapply the move");
		else pass("redo reapplied the move");
	} finally {
		pkg.dispose();
	}
}

/**
 * Drive the whole editor command set, then undo it all and check the package is
 * byte-identical to where it started.
 *
 * That round trip is the strongest single assertion available here: it proves
 * every command records enough to be reversed, including the ones that create
 * and delete parts.
 */
function checkEditorCommands(path: string): void {
	const name = basename(path);
	const { deck, pkg } = parseDeck(readFileSync(path), name);
	let model = deck;

	const selection = new Selection();
	const editor = new DeckEditor(pkg, {
		onChanged: (rebuild) => {
			if (rebuild) model = rebuildDeck(pkg, name);
		},
	});
	const ctx = (index = 0): CommandContext => ({
		editor,
		pkg,
		deck: model,
		slide: model.slides[index],
		selection,
	});
	const ownShapes = (index = 0) =>
		model.slides[index].shapes.slice(model.slides[index].templateShapes);

	// Where the package stands before any command runs.
	const original: Record<string, Uint8Array | null> = {};
	for (const part of pkg.partPaths()) original[part] = pkg.serializePart(part);
	const originalParts = new Set(pkg.partPaths());

	let steps = 0;
	const step = (label: string, run: () => boolean, expect?: () => string | null): void => {
		if (!run()) {
			fail(`${label} did nothing`);
			return;
		}
		steps++;
		const problem = expect?.();
		if (problem) fail(`${label}: ${problem}`);
		else pass(label);
	};

	try {
		const before = ownShapes().length;

		step(
			"insert text box",
			() => insertTextBox(ctx()),
			() => (ownShapes().length === before + 1 ? null : "shape count did not rise"),
		);
		const textBoxId = ownShapes()[ownShapes().length - 1].id;
		step(
			"insert shape",
			() => insertAutoShape(ctx(), "roundRect", "Rounded rectangle"),
			() => (ownShapes().length === before + 2 ? null : "shape count did not rise"),
		);
		const shapeId = ownShapes()[ownShapes().length - 1].id;
		selection.set(0, [shapeId]);
		step(
			"change shape",
			() => changeShape(ctx(), "hexagon", "Hexagon"),
			() => {
				const shape = ownShapes().find((s) => s.id === shapeId);
				return shape?.kind === "shape" && shape.geom === "hexagon"
					? null
					: "the preset did not change";
			},
		);
		step(
			"insert table",
			() => insertTable(ctx(), 3, 3),
			() => {
				const table = ownShapes().find((s) => s.kind === "table");
				if (!table || table.kind !== "table") return "no table in the model";
				return table.table.rows.length === 3 && table.table.rows[0].cells.length === 3
					? null
					: "table has the wrong shape";
			},
		);

		// Table structure, driven the way the ribbon drives it.
		const tableId = ownShapes().find((s) => s.kind === "table")!.id;
		const tableSel = new TableSelection();
		const tableOf = () => {
			const shape = ownShapes().find((s) => s.id === tableId);
			return shape?.kind === "table" ? shape.table : null;
		};
		selection.set(0, [tableId]);
		tableSel.select(tableId, 1, 1, false);

		step(
			"insert table row",
			() => insertTableRow(ctx(), tableSel, "below"),
			() => (tableOf()?.rows.length === 4 ? null : "row count did not rise"),
		);
		step(
			"insert table column",
			() => insertTableColumn(ctx(), tableSel, "right"),
			() => (tableOf()?.columns.length === 4 ? null : "column count did not rise"),
		);
		step(
			"merge cells",
			() => {
				tableSel.select(tableId, 1, 1, false);
				tableSel.select(tableId, 1, 2, true);
				return mergeTableCells(ctx(), tableSel);
			},
			() => {
				const cell = tableOf()?.rows[1]?.cells[1];
				return cell && cell.colSpan === 2 ? null : "the merge did not take";
			},
		);
		step(
			"split cells",
			() => splitTableCells(ctx(), tableSel),
			() => {
				const cell = tableOf()?.rows[1]?.cells[1];
				return cell && cell.colSpan === 1 ? null : "the split did not take";
			},
		);
		step(
			"cell fill",
			() => setCellFill(ctx(), tableSel, "#ffcc00"),
			() => {
				const cell = tableOf()?.rows[1]?.cells[1];
				return cell?.fill?.kind === "solid" && cell.fill.color === "#ffcc00"
					? null
					: "the cell fill did not come back";
			},
		);
		step(
			"delete table row",
			() => {
				tableSel.select(tableId, 3, 0, false);
				return deleteTableRows(ctx(), tableSel);
			},
			() => (tableOf()?.rows.length === 3 ? null : "row count did not fall"),
		);
		step(
			"delete table column",
			() => {
				tableSel.select(tableId, 0, 3, false);
				return deleteTableColumns(ctx(), tableSel);
			},
			() => (tableOf()?.columns.length === 3 ? null : "column count did not fall"),
		);
		step(
			"insert picture",
			() =>
				insertPicture(ctx(), {
					bytes: encodePng(64, 40, (x, y) => [x * 3, y * 5, 200]),
					extension: "png",
					name: "smoke.png",
					width: 64,
					height: 40,
				}),
			() => {
				const media = pkg.partPaths().filter((p) => p.startsWith("ppt/media/"));
				const picture = ownShapes().find((s) => s.kind === "image");
				if (!picture) return "no picture in the model";
				if (media.length === 0) return "no media part was added";
				return picture.kind === "image" && picture.url ? null : "picture has no resolvable image";
			},
		);

		// Selection-driven commands.
		const ids = ownShapes()
			.slice(-2)
			.map((s) => s.id);
		selection.set(0, ids);
		step(
			"group",
			() => groupSelection(ctx()),
			() => (ownShapes().some((s) => s.kind === "group") ? null : "no group appeared"),
		);
		step(
			"ungroup",
			() => ungroupSelection(ctx()),
			() => null,
		);

		selection.set(0, [ownShapes()[ownShapes().length - 1].id]);
		step("bring to back", () => reorderSelection(ctx(), "back"), () => null);
		step("align left", () => alignSelection(ctx(), "left"), () => null);
		step(
			"duplicate",
			() => duplicateSelection(ctx()),
			() => null,
		);
		selection.set(0, [textBoxId]);
		step(
			"bold",
			() => applyTextFormat(ctx(), { bold: true }, "Bold"),
			() => {
				const box = ownShapes().find((s) => s.id === textBoxId);
				if (!box || box.kind !== "shape" || !box.text) return "the text box vanished";
				return box.text.paragraphs.some((p) => p.runs.some((r) => r.bold))
					? null
					: "no run came back bold";
			},
		);
		step(
			"fill",
			() => setFill(ctx(), "#123456"),
			() => {
				const box = ownShapes().find((s) => s.id === textBoxId);
				return box?.kind === "shape" && box.fill?.kind === "solid" && box.fill.color === "#123456"
					? null
					: "the fill did not come back";
			},
		);
		step(
			"rotate 90",
			() => rotateBy(ctx(), 90),
			() => {
				const box = ownShapes().find((s) => s.id === textBoxId);
				return box && Math.abs(box.frame.rot - 90) < 0.5 ? null : "rotation did not stick";
			},
		);
		step(
			"flip horizontally",
			() => flipSelection(ctx(), "h"),
			() => {
				const box = ownShapes().find((s) => s.id === textBoxId);
				return box?.frame.flipH ? null : "the flip did not stick";
			},
		);
		step(
			"set position and size",
			() => setGeometry(ctx(), { x: 40, y: 60, w: 200, h: 90 }, "Set position"),
			() => {
				const box = ownShapes().find((s) => s.id === textBoxId);
				if (!box) return "the shape vanished";
				const f = box.frame;
				return Math.abs(f.x - 40) < 0.5 && Math.abs(f.w - 200) < 0.5
					? null
					: `frame came back as ${f.x.toFixed(1)},${f.y.toFixed(1)} ${f.w.toFixed(1)}x${f.h.toFixed(1)}`;
			},
		);
		step(
			"hyperlink",
			() => setHyperlink(ctx(), "https://obsidian.md"),
			() => {
				const box = ownShapes().find((s) => s.id === textBoxId);
				if (box?.kind !== "shape" || !box.text) return "the text box vanished";
				return box.text.paragraphs.some((p) => p.runs.some((r) => r.link))
					? null
					: "no run came back linked";
			},
		);
		// Copying formatting only fills a buffer, so it is deliberately not undoable
		// and must not be counted as a step.
		if (copyFormatting(ctx())) pass("copy formatting");
		else fail("copy formatting did nothing");
		step("paste formatting", () => pasteFormatting(ctx()), () => null);
		// Guides live in ppt/viewProps.xml, a part this deck may not even have yet.
		step(
			"add guides",
			() =>
				editor.transact("Guides", guideParts(), () =>
					writeGuides(pkg, [
						{ orientation: "vert", position: 640 },
						{ orientation: "horz", position: 360 },
					]),
				),
			() => {
				const guides = readGuides(pkg);
				if (guides.length !== 2) return `read back ${guides.length} guides`;
				const vert = guides.find((g) => g.orientation === "vert");
				return vert && Math.abs(vert.position - 640) < 0.5
					? null
					: `vertical guide came back at ${vert?.position}`;
			},
		);

		// Cropping moves the frame and the source rectangle together.
		const picture = ownShapes().find((s) => s.kind === "image");
		if (picture?.source) {
			const pictureId = picture.id;
			const before = { ...picture.frame };
			step(
				"crop picture",
				() =>
					editor.transact("Crop", [model.slides[0].partPath], () => {
						writeCrop(picture.source!, { l: 0.1, t: 0.2, r: 0.1, b: 0 });
						writeFrame(picture.source!, {
							...before,
							x: before.x + before.w * 0.1,
							w: before.w * 0.8,
							h: before.h * 0.8,
						});
						return true;
					}),
				() => {
					const shape = ownShapes().find((s) => s.id === pictureId);
					if (shape?.kind !== "image") return "the picture vanished";
					const crop = shape.crop;
					if (!crop) return "no crop came back";
					return Math.abs(crop.l - 0.1) < 0.001 && Math.abs(crop.t - 0.2) < 0.001
						? null
						: `crop came back as ${JSON.stringify(crop)}`;
				},
			);
		}

		const anyShape = ownShapes()[0];
		step(
			"hide shape",
			() => setShapeHidden(ctx(), anyShape.id, true),
			() => (ownShapes().find((s) => s.id === anyShape.id)?.hidden ? null : "still visible"),
		);
		step(
			"rename shape",
			() => renameShape(ctx(), anyShape.id, "Renamed by the smoke test"),
			() =>
				ownShapes().find((s) => s.id === anyShape.id)?.name === "Renamed by the smoke test"
					? null
					: "the name did not stick",
		);

		step(
			"slide background",
			() => setSlideBackgroundColor(ctx(), "#eef2f6"),
			() =>
				model.slides[0].background?.kind === "solid" &&
				model.slides[0].background.color === "#eef2f6"
					? null
					: "the background did not come back",
		);
		step(
			"delete",
			() => deleteSelection(ctx()),
			() => (ownShapes().some((s) => s.id === textBoxId) ? "the shape is still there" : null),
		);

		// Slide-level commands.
		const slidesBefore = model.slides.length;
		step(
			"new slide",
			() => newSlide(ctx()) >= 0,
			() => (model.slides.length === slidesBefore + 1 ? null : "slide count did not rise"),
		);
		step(
			"duplicate slide",
			() => duplicateCurrentSlide(ctx()) >= 0,
			() => (model.slides.length === slidesBefore + 2 ? null : "slide count did not rise"),
		);
		step(
			"reorder slides",
			() => reorderSlide(ctx(), 0, 1),
			() => null,
		);
		step(
			"delete slide",
			() => deleteCurrentSlide(ctx()),
			() => (model.slides.length === slidesBefore + 1 ? null : "slide count did not fall"),
		);

		// The edited deck must still be a deck.
		const edited = reopen(pkg, name);
		if (edited.slides.length !== model.slides.length) {
			fail(`edited deck reopened with ${edited.slides.length} slides, expected ${model.slides.length}`);
		} else {
			pass(`edited deck reopens with ${edited.slides.length} slides`);
		}
		const keep = process.env.PPTX_SMOKE_EDITED;
		if (keep) {
			writeFileSync(keep, pkg.toZip());
			console.log(`    wrote the edited deck to ${keep}`);
		}

		// Undo everything and compare against where we started.
		for (let i = 0; i < steps; i++) {
			if (!editor.undo()) {
				fail(`ran out of undo after ${i} of ${steps} steps`);
				break;
			}
		}
		const nowParts = new Set(pkg.partPaths());
		const added = [...nowParts].filter((p) => !originalParts.has(p));
		const removed = [...originalParts].filter((p) => !nowParts.has(p));
		const differing = [...originalParts].filter((p) => {
			const before = original[p];
			const after = pkg.serializePart(p);
			if (!before || !after) return true;
			return Buffer.compare(Buffer.from(before), Buffer.from(after)) !== 0;
		});

		if (added.length || removed.length || differing.length) {
			fail(
				`undo did not restore the package: ` +
					`${added.length} extra, ${removed.length} missing, ${differing.length} changed` +
					(differing.length ? ` (${differing.slice(0, 3).join(", ")})` : ""),
			);
		} else {
			pass(`undo of ${steps} commands restored every part exactly`);
		}
	} finally {
		pkg.dispose();
	}
}

// --------------------------------------------------------------------- main

const files = process.argv.slice(2);
if (files.length === 0) {
	console.error("usage: npm run smoke -- <file.pptx> [...]");
	process.exit(2);
}

// PPTX_SMOKE_HTML dumps a rendered slide as standalone HTML, so the renderer's
// output can be eyeballed against a real PowerPoint consumer.
const dumpPath = process.env.PPTX_SMOKE_HTML;
if (dumpPath) {
	const index = Number(process.env.PPTX_SMOKE_SLIDE ?? 1) - 1;
	const { deck, pkg } = parseDeck(readFileSync(files[0]), basename(files[0]));
	const slide = deck.slides[index];
	if (slide) {
		const el = renderSlide(deck, slide);
		writeFileSync(
			dumpPath,
			`<!doctype html><meta charset="utf-8"><body style="margin:0;background:#fff">` +
				`<div style="width:${deck.width}px;height:${deck.height}px">${el.outerHTML}</div>`,
		);
		console.log(`Wrote slide ${index + 1} to ${dumpPath}`);
	}
	pkg.dispose();
	process.exit(0);
}

console.log(`Parsing and rendering ${files.length} deck(s)\n`);
for (const file of files) checkDeck(file);

console.log(`\nEditing ${basename(files[0])} through the rendered DOM:`);
checkRunEdit(files[0]);
checkParagraphDelete(files[0]);

console.log(`\nCharts in ${basename(files[0])}:`);
checkCharts(files[0]);

console.log(`\nMoving and resizing shapes in ${basename(files[0])}:`);
checkGeometryEdit(files[0]);
checkUndoRedo(files[0]);

console.log(`\nEditor commands on ${basename(files[0])}:`);
checkEditorCommands(files[0]);

console.log(failures === 0 ? "\nAll checks passed." : `\n${failures} check(s) failed.`);
process.exit(failures === 0 ? 0 : 1);
