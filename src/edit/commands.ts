import type { PptxPackage } from "../pptx/package";
import type { Deck, Frame, Shape, Slide } from "../pptx/types";
import { child } from "../pptx/xml";
import {
	findShapesById,
	groupChildren,
	insertAfter,
	nextShapeId,
	p as pEl,
	a as aEl,
	reassignIds,
	remapRelationships,
	removeShape,
	shapeElements,
	shapeIdOf,
	spTreeOf,
} from "../ooxml/tree";
import { relsPathFor } from "../ooxml/rels";
import type { DeckEditor } from "./DeckEditor";
import { readFrame, writeFrame } from "./geometryWrite";
import type { Selection } from "./Selection";
import { EMU_PER_PX } from "../pptx/types";

export interface CommandContext {
	editor: DeckEditor;
	pkg: PptxPackage;
	deck: Deck;
	slide: Slide;
	selection: Selection;
}

/** The selected shapes, in z-order, with their resolved frames. */
export function selectedShapes(ctx: CommandContext): Shape[] {
	const own = ctx.slide.shapes.slice(ctx.slide.templateShapes);
	return own.filter((s) => s.source && ctx.selection.has(s.id));
}

function slidePath(ctx: CommandContext): string {
	return ctx.slide.partPath;
}

/** Parts a slide edit may touch: the slide and its relationships. */
function slideParts(ctx: CommandContext): string[] {
	return [slidePath(ctx), relsPathFor(slidePath(ctx))];
}

// -------------------------------------------------------------- delete

export function deleteSelection(ctx: CommandContext): boolean {
	const path = slidePath(ctx);
	const tree = spTreeOf(ctx.pkg, path);
	const targets = findShapesById(tree, new Set(ctx.selection.ids));
	if (targets.length === 0) return false;

	const done = ctx.editor.transact(
		targets.length > 1 ? `Delete ${targets.length} shapes` : "Delete shape",
		[path],
		() => {
			for (const el of targets) removeShape(el);
			return true;
		},
	);
	if (done) ctx.selection.clear();
	return done;
}

// ------------------------------------------------------------ clipboard

interface ClipboardEntry {
	xml: string;
	part: string;
}

/** Shapes live here between a copy and a paste, as serialised XML. */
let clipboard: ClipboardEntry[] = [];

export function hasClipboard(): boolean {
	return clipboard.length > 0;
}

export function copySelection(ctx: CommandContext): boolean {
	const tree = spTreeOf(ctx.pkg, slidePath(ctx));
	const targets = findShapesById(tree, new Set(ctx.selection.ids));
	if (targets.length === 0) return false;
	const serializer = new XMLSerializer();
	clipboard = targets.map((el) => ({ xml: serializer.serializeToString(el), part: slidePath(ctx) }));
	return true;
}

export function cutSelection(ctx: CommandContext): boolean {
	if (!copySelection(ctx)) return false;
	return deleteSelection(ctx);
}

export function pasteClipboard(ctx: CommandContext, offset = 12): boolean {
	if (clipboard.length === 0) return false;
	const path = slidePath(ctx);
	const tree = spTreeOf(ctx.pkg, path);
	if (!tree) return false;

	const pasted: string[] = [];
	const done = ctx.editor.transact("Paste", slideParts(ctx), () => {
		const doc = tree.ownerDocument;
		let counter = nextShapeId(tree);
		const next = () => counter++;

		for (const entry of clipboard) {
			const parsed = new DOMParser().parseFromString(entry.xml, "application/xml");
			if (parsed.getElementsByTagName("parsererror")[0]) continue;
			const el = doc.importNode(parsed.documentElement, true) as Element;
			reassignIds(el, next);
			remapRelationships(ctx.pkg, el, entry.part, path);
			const frame = readFrame(el);
			if (frame) writeFrame(el, { ...frame, x: frame.x + offset, y: frame.y + offset });
			tree.appendChild(el);
			const id = shapeIdOf(el);
			if (id) pasted.push(id);
		}
		return pasted.length > 0;
	});

	if (done) ctx.selection.set(ctx.slide.index - 1, pasted);
	return done;
}

export function duplicateSelection(ctx: CommandContext): boolean {
	const saved = clipboard;
	try {
		if (!copySelection(ctx)) return false;
		return pasteClipboard(ctx, 12);
	} finally {
		// Duplicating should not clobber whatever the user actually copied.
		clipboard = saved.length > 0 && saved !== clipboard ? saved : clipboard;
	}
}

// -------------------------------------------------------------- z-order

export type ReorderMode = "front" | "back" | "forward" | "backward";

export function reorderSelection(ctx: CommandContext, mode: ReorderMode): boolean {
	const path = slidePath(ctx);
	const tree = spTreeOf(ctx.pkg, path);
	if (!tree) return false;
	const ids = new Set(ctx.selection.ids);
	const targets = findShapesById(tree, ids);
	if (targets.length === 0) return false;

	const labels: Record<ReorderMode, string> = {
		front: "Bring to front",
		back: "Send to back",
		forward: "Bring forward",
		backward: "Send backward",
	};

	return ctx.editor.transact(labels[mode], [path], () => {
		const order = shapeElements(tree);
		if (mode === "front") {
			for (const el of targets) tree.appendChild(el);
			return true;
		}
		if (mode === "back") {
			const first = order[0];
			let anchor: Node | null = first ?? null;
			for (const el of targets) tree.insertBefore(el, anchor);
			// Keep relative order among the moved shapes.
			anchor = targets[targets.length - 1]?.nextSibling ?? null;
			return true;
		}
		// One step: walk from the end for "forward" so shapes cannot pass each other.
		const sequence = mode === "forward" ? [...targets].reverse() : targets;
		for (const el of sequence) {
			const siblings = shapeElements(tree);
			const index = siblings.indexOf(el);
			if (index === -1) continue;
			if (mode === "forward") {
				const next = siblings[index + 1];
				if (next && !ids.has(shapeIdOf(next) ?? "")) insertAfter(tree, el, next);
			} else {
				const previous = siblings[index - 1];
				if (previous && !ids.has(shapeIdOf(previous) ?? "")) tree.insertBefore(el, previous);
			}
		}
		return true;
	});
}

// ---------------------------------------------------------------- align

export type AlignMode = "left" | "centerH" | "right" | "top" | "middle" | "bottom";

export function alignSelection(ctx: CommandContext, mode: AlignMode): boolean {
	const shapes = selectedShapes(ctx);
	if (shapes.length === 0) return false;

	// One shape aligns to the slide; several align to each other, which is what
	// PowerPoint does and what people expect from a single Ctrl-click.
	const bounds =
		shapes.length === 1
			? { x: 0, y: 0, w: ctx.deck.width, h: ctx.deck.height }
			: unionBounds(shapes.map((s) => s.frame));

	const labels: Record<AlignMode, string> = {
		left: "Align left",
		centerH: "Align centre",
		right: "Align right",
		top: "Align top",
		middle: "Align middle",
		bottom: "Align bottom",
	};

	return ctx.editor.transact(labels[mode], [slidePath(ctx)], () => {
		let changed = false;
		for (const shape of shapes) {
			if (!shape.source) continue;
			const f = shape.frame;
			const next = { ...f };
			switch (mode) {
				case "left":
					next.x = bounds.x;
					break;
				case "centerH":
					next.x = bounds.x + (bounds.w - f.w) / 2;
					break;
				case "right":
					next.x = bounds.x + bounds.w - f.w;
					break;
				case "top":
					next.y = bounds.y;
					break;
				case "middle":
					next.y = bounds.y + (bounds.h - f.h) / 2;
					break;
				case "bottom":
					next.y = bounds.y + bounds.h - f.h;
					break;
			}
			if (Math.abs(next.x - f.x) > 0.01 || Math.abs(next.y - f.y) > 0.01) {
				writeFrame(shape.source, next);
				changed = true;
			}
		}
		return changed;
	});
}

export function distributeSelection(ctx: CommandContext, axis: "h" | "v"): boolean {
	const shapes = selectedShapes(ctx);
	if (shapes.length < 3) return false;

	const sorted = [...shapes].sort((a, b) =>
		axis === "h" ? a.frame.x - b.frame.x : a.frame.y - b.frame.y,
	);
	const first = sorted[0].frame;
	const last = sorted[sorted.length - 1].frame;
	const span =
		axis === "h" ? last.x + last.w - first.x : last.y + last.h - first.y;
	const used = sorted.reduce((total, s) => total + (axis === "h" ? s.frame.w : s.frame.h), 0);
	const gap = (span - used) / (sorted.length - 1);

	return ctx.editor.transact(
		axis === "h" ? "Distribute horizontally" : "Distribute vertically",
		[slidePath(ctx)],
		() => {
			let cursor = axis === "h" ? first.x : first.y;
			let changed = false;
			for (const shape of sorted) {
				if (!shape.source) continue;
				const f = shape.frame;
				const next = axis === "h" ? { ...f, x: cursor } : { ...f, y: cursor };
				if (Math.abs((axis === "h" ? next.x : next.y) - (axis === "h" ? f.x : f.y)) > 0.01) {
					writeFrame(shape.source, next);
					changed = true;
				}
				cursor += (axis === "h" ? f.w : f.h) + gap;
			}
			return changed;
		},
	);
}

// ---------------------------------------------------------------- group

export function groupSelection(ctx: CommandContext): boolean {
	const shapes = selectedShapes(ctx);
	if (shapes.length < 2) return false;
	const path = slidePath(ctx);
	const tree = spTreeOf(ctx.pkg, path);
	if (!tree) return false;

	const bounds = unionBounds(shapes.map((s) => s.frame));
	let groupId = "";

	const done = ctx.editor.transact("Group", [path], () => {
		const doc = tree.ownerDocument;
		const id = nextShapeId(tree);
		groupId = String(id);

		const group = pEl(doc, "grpSp");
		const nvGrpSpPr = pEl(doc, "nvGrpSpPr");
		const cNvPr = pEl(doc, "cNvPr");
		cNvPr.setAttribute("id", groupId);
		cNvPr.setAttribute("name", `Group ${groupId}`);
		nvGrpSpPr.appendChild(cNvPr);
		nvGrpSpPr.appendChild(pEl(doc, "cNvGrpSpPr"));
		nvGrpSpPr.appendChild(pEl(doc, "nvPr"));
		group.appendChild(nvGrpSpPr);

		const grpSpPr = pEl(doc, "grpSpPr");
		const xfrm = aEl(doc, "xfrm");
		const emu = (v: number) => String(Math.round(v * EMU_PER_PX));
		const off = aEl(doc, "off");
		off.setAttribute("x", emu(bounds.x));
		off.setAttribute("y", emu(bounds.y));
		const ext = aEl(doc, "ext");
		ext.setAttribute("cx", emu(bounds.w));
		ext.setAttribute("cy", emu(bounds.h));
		// Child space matches the group's own box, so members keep the coordinates
		// they already had and nothing shifts at the moment of grouping.
		const chOff = aEl(doc, "chOff");
		chOff.setAttribute("x", emu(bounds.x));
		chOff.setAttribute("y", emu(bounds.y));
		const chExt = aEl(doc, "chExt");
		chExt.setAttribute("cx", emu(bounds.w));
		chExt.setAttribute("cy", emu(bounds.h));
		xfrm.append(off, ext, chOff, chExt);
		grpSpPr.appendChild(xfrm);
		group.appendChild(grpSpPr);

		const topmost = shapes[shapes.length - 1].source!;
		tree.insertBefore(group, topmost.nextSibling);

		for (const shape of shapes) {
			if (!shape.source) continue;
			// A placeholder inside a group must carry its own frame; otherwise it
			// would keep inheriting one from a layout that knows nothing about it.
			writeFrame(shape.source, shape.frame);
			group.appendChild(shape.source);
		}
		return true;
	});

	if (done && groupId) ctx.selection.set(ctx.slide.index - 1, [groupId]);
	return done;
}

export function ungroupSelection(ctx: CommandContext): boolean {
	const shapes = selectedShapes(ctx).filter((s) => s.kind === "group");
	if (shapes.length === 0) return false;
	const path = slidePath(ctx);
	const tree = spTreeOf(ctx.pkg, path);
	if (!tree) return false;

	const freed: string[] = [];
	const done = ctx.editor.transact("Ungroup", [path], () => {
		for (const shape of shapes) {
			if (shape.kind !== "group" || !shape.source) continue;
			const group = shape.source;
			const sx = shape.childOffset.w > 0 ? shape.frame.w / shape.childOffset.w : 1;
			const sy = shape.childOffset.h > 0 ? shape.frame.h / shape.childOffset.h : 1;

			for (const kid of shape.children) {
				if (!kid.source) continue;
				// Map the child out of the group's coordinate space and back onto the slide.
				writeFrame(kid.source, {
					x: shape.frame.x + (kid.frame.x - shape.childOffset.x) * sx,
					y: shape.frame.y + (kid.frame.y - shape.childOffset.y) * sy,
					w: kid.frame.w * sx,
					h: kid.frame.h * sy,
				});
				tree.insertBefore(kid.source, group);
				const id = shapeIdOf(kid.source);
				if (id) freed.push(id);
			}
			// Anything the model did not surface (an unsupported child) still has to
			// come out, or ungrouping would delete it along with the group.
			for (const orphan of groupChildren(group)) tree.insertBefore(orphan, group);
			removeShape(group);
		}
		return freed.length > 0;
	});

	if (done) ctx.selection.set(ctx.slide.index - 1, freed);
	return done;
}

// -------------------------------------------------------------- helpers

export function unionBounds(frames: Frame[]): { x: number; y: number; w: number; h: number } {
	const left = Math.min(...frames.map((f) => f.x));
	const top = Math.min(...frames.map((f) => f.y));
	const right = Math.max(...frames.map((f) => f.x + f.w));
	const bottom = Math.max(...frames.map((f) => f.y + f.h));
	return { x: left, y: top, w: right - left, h: bottom - top };
}

/** True when a shape can be ungrouped. */
export function isGroup(shape: Shape): boolean {
	return shape.kind === "group" && shape.source !== null && child(shape.source, "grpSpPr") !== null;
}
