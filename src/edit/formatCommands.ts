import type { Run, Shape, TextBody } from "../pptx/types";
import { PT_TO_PX } from "../pptx/types";
import { child } from "../pptx/xml";
import {
	type OutlinePatch,
	type ParagraphPatch,
	type RunPatch,
	applyParagraphPatch,
	applyParagraphPatchToBody,
	applyRunPatch,
	applyRunPatchToBody,
	applyShapeFill,
	applyShapeOutline,
	applyTextAnchor,
	splitRun,
} from "../ooxml/format";
import { relsPathFor } from "../ooxml/rels";
import type { CommandContext } from "./commands";
import { selectedShapes } from "./commands";
import { currentTextSelection } from "./textSelection";

function parts(ctx: CommandContext): string[] {
	return [ctx.slide.partPath, relsPathFor(ctx.slide.partPath)];
}

/** Text bodies the ribbon should act on: the whole shape, when nothing is selected. */
function targetBodies(ctx: CommandContext): { shape: Shape; body: TextBody }[] {
	const out: { shape: Shape; body: TextBody }[] = [];
	for (const shape of selectedShapes(ctx)) {
		if (shape.kind === "shape" && shape.text?.source) out.push({ shape, body: shape.text });
		if (shape.kind === "table") {
			for (const row of shape.table.rows) {
				for (const cell of row.cells) {
					if (cell.text?.source) out.push({ shape, body: cell.text });
				}
			}
		}
	}
	return out;
}

/**
 * Apply character formatting.
 *
 * With text selected inside an open editor the patch lands on exactly those
 * characters, splitting runs where it has to. Otherwise it lands on every run in
 * the selected shapes, which is what a ribbon button means when the shape rather
 * than the text is selected.
 */
export function applyTextFormat(ctx: CommandContext, patch: RunPatch, label: string): boolean {
	const selection = currentTextSelection();

	if (selection && selection.runs.length > 0) {
		return ctx.editor.transact(label, parts(ctx), () => {
			for (const { run, start, end } of selection.runs) {
				if (!run.source) continue;
				const target = splitRun(run.source, start, end);
				if (target) applyRunPatch(target, patch);
			}
			return true;
		});
	}

	const bodies = targetBodies(ctx);
	if (bodies.length === 0) return false;
	return ctx.editor.transact(label, parts(ctx), () => {
		let changed = false;
		for (const { body } of bodies) {
			if (body.source && applyRunPatchToBody(body.source, patch)) changed = true;
		}
		return changed;
	});
}

export function applyParagraphFormat(
	ctx: CommandContext,
	patch: ParagraphPatch,
	label: string,
): boolean {
	const selection = currentTextSelection();

	if (selection && selection.paragraphs.length > 0) {
		return ctx.editor.transact(label, parts(ctx), () => {
			for (const para of selection.paragraphs) applyParagraphPatch(para, patch);
			return true;
		});
	}

	const bodies = targetBodies(ctx);
	if (bodies.length === 0) return false;
	return ctx.editor.transact(label, parts(ctx), () => {
		let changed = false;
		for (const { body } of bodies) {
			if (body.source && applyParagraphPatchToBody(body.source, patch)) changed = true;
		}
		return changed;
	});
}

export function setFill(ctx: CommandContext, color: string | null): boolean {
	const shapes = selectedShapes(ctx).filter((s) => s.source && child(s.source, "spPr"));
	if (shapes.length === 0) return false;
	return ctx.editor.transact(color === null ? "No fill" : "Fill", parts(ctx), () => {
		let changed = false;
		for (const shape of shapes) {
			if (shape.source && applyShapeFill(shape.source, color)) changed = true;
		}
		return changed;
	});
}

export function setOutline(ctx: CommandContext, patch: OutlinePatch, label: string): boolean {
	const shapes = selectedShapes(ctx).filter((s) => s.source && child(s.source, "spPr"));
	if (shapes.length === 0) return false;
	return ctx.editor.transact(label, parts(ctx), () => {
		let changed = false;
		for (const shape of shapes) {
			if (shape.source && applyShapeOutline(shape.source, patch)) changed = true;
		}
		return changed;
	});
}

export function setVerticalAnchor(ctx: CommandContext, anchor: "t" | "ctr" | "b"): boolean {
	const shapes = selectedShapes(ctx).filter((s) => s.kind === "shape" && s.text);
	if (shapes.length === 0) return false;
	return ctx.editor.transact("Vertical alignment", parts(ctx), () => {
		let changed = false;
		for (const shape of shapes) {
			if (shape.source && applyTextAnchor(shape.source, anchor)) changed = true;
		}
		return changed;
	});
}

// ------------------------------------------------------- current values

export interface TextState {
	bold: boolean;
	italic: boolean;
	underline: boolean;
	strike: boolean;
	/** Points, or null when the selection is mixed. */
	size: number | null;
	color: string | null;
	font: string | null;
	align: string | null;
	bulleted: boolean;
}

/** What the ribbon should show for the current selection. */
export function textState(ctx: CommandContext | null): TextState | null {
	if (!ctx) return null;
	const runs: Run[] = [];
	let align: string | null = null;
	let bulleted = false;

	for (const shape of selectedShapes(ctx)) {
		const bodies: TextBody[] = [];
		if (shape.kind === "shape" && shape.text) bodies.push(shape.text);
		if (shape.kind === "table") {
			for (const row of shape.table.rows) {
				for (const cell of row.cells) if (cell.text) bodies.push(cell.text);
			}
		}
		for (const body of bodies) {
			for (const para of body.paragraphs) {
				align ??= para.align;
				if (para.bullet) bulleted = true;
				for (const run of para.runs) if (run.text !== "") runs.push(run);
			}
		}
	}
	if (runs.length === 0) return null;

	const all = <T>(pick: (r: Run) => T): T | null => {
		const first = pick(runs[0]);
		return runs.every((r) => pick(r) === first) ? first : null;
	};

	return {
		bold: runs.every((r) => r.bold),
		italic: runs.every((r) => r.italic),
		underline: runs.every((r) => r.underline),
		strike: runs.every((r) => r.strike),
		size: (() => {
			const px = all((r) => Math.round(r.size * 100) / 100);
			return px === null ? null : Math.round((px / PT_TO_PX) * 10) / 10;
		})(),
		color: all((r) => r.color),
		font: all((r) => r.font),
		align,
		bulleted,
	};
}

/** The fill colour shown on the ribbon, when the selection agrees on one. */
export function fillState(ctx: CommandContext | null): string | null {
	if (!ctx) return null;
	const shapes = selectedShapes(ctx);
	if (shapes.length === 0) return null;
	const first = shapes[0];
	if (first.kind !== "shape" || !first.fill || first.fill.kind !== "solid") return null;
	return first.fill.color;
}

export function outlineState(ctx: CommandContext | null): string | null {
	if (!ctx) return null;
	const shapes = selectedShapes(ctx);
	const first = shapes[0];
	if (!first) return null;
	if (first.kind === "shape" || first.kind === "line") return first.stroke?.color ?? null;
	return null;
}
