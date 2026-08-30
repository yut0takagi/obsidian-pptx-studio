import { CONTENT_TYPES_PART } from "../ooxml/contentTypes";
import { relsPathFor } from "../ooxml/rels";
import {
	PRESENTATION,
	addSlide,
	availableLayouts,
	deleteSlide,
	duplicateSlide,
	layoutOf,
	moveSlide,
	nextSlidePath,
	slidePaths,
} from "../ooxml/slideOps";
import type { CommandContext } from "./commands";

/**
 * Slide-level edits touch the presentation, its relationships and the content
 * type map as well as the slide parts themselves. Listing them all here means
 * one undo step puts every one of them back together.
 */
function parts(ctx: CommandContext, extra: string[] = []): string[] {
	return [PRESENTATION, relsPathFor(PRESENTATION), CONTENT_TYPES_PART, ...extra];
}

export function listLayouts(ctx: CommandContext): { path: string; name: string }[] {
	return availableLayouts(ctx.pkg);
}

/** Insert a slide after the current one. Returns its position, or -1. */
export function newSlide(ctx: CommandContext, layoutPath?: string): number {
	const position = ctx.slide.index; // 1-based index of the current slide == insert after it
	const layout = layoutPath ?? layoutOf(ctx.pkg, ctx.slide.partPath) ?? availableLayouts(ctx.pkg)[0]?.path;
	if (!layout) return -1;

	// The path is predictable, so the part being created can be snapshotted along
	// with the rest: one undo then puts the whole operation back, not two.
	const created = nextSlidePath(ctx.pkg);
	const done = ctx.editor.transact(
		"New slide",
		parts(ctx, [created, relsPathFor(created)]),
		() => {
			addSlide(ctx.pkg, layout, position, created);
			return true;
		},
	);
	return done ? position : -1;
}

export function duplicateCurrentSlide(ctx: CommandContext): number {
	const position = ctx.slide.index;
	const source = ctx.slide.partPath;
	const created = nextSlidePath(ctx.pkg);

	const done = ctx.editor.transact(
		"Duplicate slide",
		parts(ctx, [source, created, relsPathFor(created)]),
		() => {
			duplicateSlide(ctx.pkg, source, position, created);
			return true;
		},
	);
	return done ? position : -1;
}

export function deleteCurrentSlide(ctx: CommandContext): boolean {
	const path = ctx.slide.partPath;
	const notes = ctx.pkg.relByKind(path, "notesSlide")?.target;
	const touched = [path, relsPathFor(path)];
	if (notes) touched.push(notes, relsPathFor(notes));

	return ctx.editor.transact("Delete slide", parts(ctx, touched), () =>
		deleteSlide(ctx.pkg, path),
	);
}

/** Move the current slide by `delta` positions. Returns its new index, or -1. */
export function moveCurrentSlide(ctx: CommandContext, delta: number): number {
	const order = slidePaths(ctx.pkg);
	const from = order.indexOf(ctx.slide.partPath);
	if (from === -1) return -1;
	const to = Math.max(0, Math.min(order.length - 1, from + delta));
	if (to === from) return -1;

	const done = ctx.editor.transact("Reorder slides", parts(ctx), () =>
		moveSlide(ctx.pkg, from, to),
	);
	return done ? to : -1;
}

/** Move a slide from one position to another, for thumbnail drag and drop. */
export function reorderSlide(ctx: CommandContext, from: number, to: number): boolean {
	return ctx.editor.transact("Reorder slides", parts(ctx), () => moveSlide(ctx.pkg, from, to));
}

export function canDeleteSlide(ctx: CommandContext): boolean {
	return ctx.deck.slides.length > 1;
}
