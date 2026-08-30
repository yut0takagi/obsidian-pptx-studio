import type { Run, Shape, TextBody } from "../pptx/types";
import { PT_TO_PX } from "../pptx/types";
import { child } from "../pptx/xml";
import {
	type OutlinePatch,
	type ParagraphPatch,
	type RunPatch,
	applyGeometry,
	applyHyperlink,
	applyParagraphPatch,
	applyParagraphPatchToBody,
	applyRunPatch,
	applyRunPatchToBody,
	applyShapeFill,
	applyShapeOutline,
	applyTextAnchor,
	splitRun,
} from "../ooxml/format";
import { REL_BASE, ensureRelationship, relsPathFor } from "../ooxml/rels";
import { writeFlip, writeFrame, writeRotation } from "./geometryWrite";
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

/**
 * Apply a paragraph patch to one paragraph of one shape.
 *
 * Tab arrives while the text editor is open, and committing that edit
 * invalidates every element reference the caller was holding — so the target is
 * named by id and index and looked up afresh.
 */
export function applyParagraphFormatAt(
	ctx: CommandContext,
	shapeId: string,
	paragraphIndex: number,
	patch: ParagraphPatch,
	label: string,
): boolean {
	const own = ctx.slide.shapes.slice(ctx.slide.templateShapes);
	const shape = own.find((s) => s.id === shapeId);
	if (!shape || shape.kind !== "shape") return false;
	const para = shape.text?.paragraphs[paragraphIndex]?.source;
	if (!para) return false;
	return ctx.editor.transact(label, parts(ctx), () => {
		applyParagraphPatch(para, patch);
		return true;
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

// ------------------------------------------------------ geometry & shape

export interface GeometryPatch {
	x?: number;
	y?: number;
	w?: number;
	h?: number;
	rotation?: number;
}

/** Position, size and rotation as the ribbon's number fields show them. */
export function geometryState(ctx: CommandContext | null): {
	x: number | null;
	y: number | null;
	w: number | null;
	h: number | null;
	rotation: number | null;
} | null {
	if (!ctx) return null;
	const shapes = selectedShapes(ctx);
	if (shapes.length === 0) return null;
	const agree = (pick: (s: Shape) => number): number | null => {
		const first = Math.round(pick(shapes[0]) * 100) / 100;
		return shapes.every((s) => Math.round(pick(s) * 100) / 100 === first) ? first : null;
	};
	return {
		x: agree((s) => s.frame.x),
		y: agree((s) => s.frame.y),
		w: agree((s) => s.frame.w),
		h: agree((s) => s.frame.h),
		rotation: agree((s) => s.frame.rot),
	};
}

export function setGeometry(ctx: CommandContext, patch: GeometryPatch, label: string): boolean {
	const shapes = selectedShapes(ctx).filter((s) => s.source);
	if (shapes.length === 0) return false;
	return ctx.editor.transact(label, parts(ctx), () => {
		let changed = false;
		for (const shape of shapes) {
			if (!shape.source) continue;
			const frame = {
				...shape.frame,
				x: patch.x ?? shape.frame.x,
				y: patch.y ?? shape.frame.y,
				w: Math.max(1, patch.w ?? shape.frame.w),
				h: Math.max(1, patch.h ?? shape.frame.h),
			};
			if (writeFrame(shape.source, frame)) changed = true;
			if (patch.rotation !== undefined && writeRotation(shape.source, patch.rotation)) {
				changed = true;
			}
		}
		return changed;
	});
}

/** Rotate by a relative amount, for the 90-degree buttons. */
export function rotateBy(ctx: CommandContext, degrees: number): boolean {
	const shapes = selectedShapes(ctx).filter((s) => s.source);
	if (shapes.length === 0) return false;
	return ctx.editor.transact(`Rotate ${degrees}°`, parts(ctx), () => {
		let changed = false;
		for (const shape of shapes) {
			if (!shape.source) continue;
			// A placeholder with no a:xfrm has nowhere to record a rotation yet.
			writeFrame(shape.source, shape.frame);
			if (writeRotation(shape.source, shape.frame.rot + degrees)) changed = true;
		}
		return changed;
	});
}

export function flipSelection(ctx: CommandContext, axis: "h" | "v"): boolean {
	const shapes = selectedShapes(ctx).filter((s) => s.source);
	if (shapes.length === 0) return false;
	return ctx.editor.transact(axis === "h" ? "Flip horizontal" : "Flip vertical", parts(ctx), () => {
		let changed = false;
		for (const shape of shapes) {
			if (!shape.source) continue;
			writeFrame(shape.source, shape.frame);
			const current = axis === "h" ? shape.frame.flipH : shape.frame.flipV;
			if (writeFlip(shape.source, axis, !current)) changed = true;
		}
		return changed;
	});
}

/** Swap the preset geometry of every selected shape. */
export function changeShape(ctx: CommandContext, preset: string, label: string): boolean {
	const shapes = selectedShapes(ctx).filter((s) => s.kind === "shape" && s.source);
	if (shapes.length === 0) return false;
	return ctx.editor.transact(`Change to ${label.toLowerCase()}`, parts(ctx), () => {
		let changed = false;
		for (const shape of shapes) {
			if (shape.source && applyGeometry(shape.source, preset)) changed = true;
		}
		return changed;
	});
}

// ----------------------------------------------------------- hyperlinks

/** Attach a hyperlink to the selected text, or to all text in the shape. */
export function setHyperlink(ctx: CommandContext, url: string | null): boolean {
	const selection = currentTextSelection();
	const targets: Element[] = [];

	if (selection && selection.runs.length > 0) {
		for (const { run, start, end } of selection.runs) {
			if (!run.source) continue;
			const split = splitRun(run.source, start, end);
			if (split) targets.push(split);
		}
	} else {
		for (const { body } of targetBodies(ctx)) {
			if (!body.source) continue;
			for (const para of Array.from(body.source.getElementsByTagName("*"))) {
				if (para.localName === "r") targets.push(para);
			}
		}
	}
	if (targets.length === 0) return false;

	return ctx.editor.transact(url ? "Add link" : "Remove link", parts(ctx), () => {
		const relId =
			url === null
				? null
				: ensureRelationship(
						ctx.pkg,
						ctx.slide.partPath,
						`${REL_BASE}/hyperlink`,
						url,
						true,
					);
		for (const runEl of targets) applyHyperlink(runEl, relId);
		return true;
	});
}

/** The link on the first selected run, so the ribbon can prefill the prompt. */
export function hyperlinkState(ctx: CommandContext | null): string | null {
	if (!ctx) return null;
	for (const shape of selectedShapes(ctx)) {
		if (shape.kind !== "shape" || !shape.text) continue;
		for (const para of shape.text.paragraphs) {
			for (const run of para.runs) if (run.link) return run.link;
		}
	}
	return null;
}

// ------------------------------------------------------- format painter

interface CopiedFormat {
	run: RunPatch;
	fill: string | null | undefined;
	outline: string | null | undefined;
}

let copiedFormat: CopiedFormat | null = null;

export function hasCopiedFormat(): boolean {
	return copiedFormat !== null;
}

/** Remember the formatting of the first selected shape. */
export function copyFormatting(ctx: CommandContext): boolean {
	const shapes = selectedShapes(ctx);
	const first = shapes[0];
	if (!first) return false;
	const text = textState(ctx);
	copiedFormat = {
		run: text
			? {
					bold: text.bold,
					italic: text.italic,
					underline: text.underline,
					strike: text.strike,
					...(text.size !== null ? { size: text.size } : {}),
					...(text.color !== null ? { color: text.color } : {}),
					...(text.font !== null ? { font: text.font } : {}),
				}
			: {},
		fill: first.kind === "shape" && first.fill?.kind === "solid" ? first.fill.color : undefined,
		outline:
			first.kind === "shape" || first.kind === "line" ? (first.stroke?.color ?? null) : undefined,
	};
	return true;
}

/** Apply the remembered formatting to the current selection. */
export function pasteFormatting(ctx: CommandContext): boolean {
	const format = copiedFormat;
	if (!format) return false;
	const shapes = selectedShapes(ctx).filter((s) => s.source);
	if (shapes.length === 0) return false;

	return ctx.editor.transact("Paste formatting", parts(ctx), () => {
		let changed = false;
		for (const shape of shapes) {
			if (!shape.source) continue;
			if (shape.kind === "shape" && shape.text?.source) {
				applyRunPatchToBody(shape.text.source, format.run);
				changed = true;
			}
			if (format.fill !== undefined && applyShapeFill(shape.source, format.fill)) changed = true;
			if (format.outline !== undefined) {
				applyShapeOutline(shape.source, { color: format.outline });
				changed = true;
			}
		}
		return changed;
	});
}
