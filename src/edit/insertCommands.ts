import { CONTENT_TYPES_PART } from "../ooxml/contentTypes";
import {
	buildLine,
	buildPicture,
	buildShape,
	buildTable,
	buildTextBox,
	type NewFrame,
	nextMediaPath,
} from "../ooxml/insert";
import { relsPathFor } from "../ooxml/rels";
import { shapeIdOf, spTreeOf } from "../ooxml/tree";
import type { CommandContext } from "./commands";

/** Where a newly inserted object lands: centred, at a comfortable default size. */
function defaultFrame(ctx: CommandContext, widthRatio: number, heightRatio: number): NewFrame {
	const w = ctx.deck.width * widthRatio;
	const h = ctx.deck.height * heightRatio;
	return { x: (ctx.deck.width - w) / 2, y: (ctx.deck.height - h) / 2, w, h };
}

function parts(ctx: CommandContext): string[] {
	return [ctx.slide.partPath, relsPathFor(ctx.slide.partPath), CONTENT_TYPES_PART];
}

/** Insert an element built by `build`, then select it. */
function insert(
	ctx: CommandContext,
	label: string,
	build: (spTree: Element) => Element | null,
	extraParts: string[] = [],
): boolean {
	const tree = spTreeOf(ctx.pkg, ctx.slide.partPath);
	if (!tree) return false;
	let newId: string | null = null;

	const done = ctx.editor.transact(label, [...parts(ctx), ...extraParts], () => {
		const el = build(tree);
		if (!el) return false;
		tree.appendChild(el);
		newId = shapeIdOf(el);
		return true;
	});

	if (done && newId) ctx.selection.set(ctx.slide.index - 1, [newId]);
	return done;
}

export function insertAutoShape(ctx: CommandContext, preset: string, label: string): boolean {
	const frame = defaultFrame(ctx, 0.25, 0.22);
	return insert(ctx, `Insert ${label.toLowerCase()}`, (tree) =>
		buildShape(tree, preset, frame, label),
	);
}

export function insertTextBox(ctx: CommandContext): boolean {
	const frame = defaultFrame(ctx, 0.4, 0.12);
	return insert(ctx, "Insert text box", (tree) => buildTextBox(tree, frame, "Text"));
}

export function insertLine(ctx: CommandContext): boolean {
	const frame = defaultFrame(ctx, 0.3, 0);
	return insert(ctx, "Insert line", (tree) => buildLine(tree, { ...frame, h: 0 }));
}

export interface InsertImage {
	bytes: Uint8Array;
	extension: string;
	name: string;
	/** Natural size in pixels, used to keep the aspect ratio. */
	width?: number;
	height?: number;
}

export function insertPicture(ctx: CommandContext, image: InsertImage): boolean {
	// Fit inside half the slide while keeping the image's own proportions.
	const maxW = ctx.deck.width * 0.5;
	const maxH = ctx.deck.height * 0.5;
	const naturalW = image.width && image.width > 0 ? image.width : maxW;
	const naturalH = image.height && image.height > 0 ? image.height : maxH;
	const scale = Math.min(maxW / naturalW, maxH / naturalH, 1);
	const w = naturalW * scale;
	const h = naturalH * scale;
	const frame: NewFrame = {
		x: (ctx.deck.width - w) / 2,
		y: (ctx.deck.height - h) / 2,
		w,
		h,
	};

	// The media part is named before the transaction starts so it is snapshotted
	// alongside everything else; otherwise undo would leave the image orphaned.
	const mediaPath = nextMediaPath(ctx.pkg, image.extension);
	return insert(
		ctx,
		"Insert picture",
		(tree) => buildPicture(ctx.pkg, tree, ctx.slide.partPath, image, frame, mediaPath),
		[mediaPath],
	);
}

export function insertTable(ctx: CommandContext, rows: number, columns: number): boolean {
	const frame = defaultFrame(ctx, 0.7, Math.min(0.6, 0.09 * rows));
	return insert(ctx, "Insert table", (tree) => buildTable(tree, frame, rows, columns));
}
