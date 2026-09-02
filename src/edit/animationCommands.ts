/**
 * Transitions and animations as undoable commands.
 *
 * Everything here touches one slide part and nothing else — no relationships,
 * no media — so each is the cheapest kind of edit there is, and rebuilding the
 * model afterwards is only needed because the timing tree is not part of it.
 */
import {
	type EffectKind,
	type Transition,
	type Trigger,
	addAnimation,
	canReorder,
	clearAnimations,
	moveAnimation,
	readAnimations,
	readTransition,
	removeAnimationsFor,
	writeTransition,
	type AnimationEntry,
} from "../ooxml/animation";
import type { CommandContext } from "./commands";
import { selectedShapes } from "./commands";

function slideRoot(ctx: CommandContext): Element | null {
	return ctx.pkg.xml(ctx.slide.partPath)?.documentElement ?? null;
}

export function slideTransition(ctx: CommandContext | null): Transition | null {
	if (!ctx) return null;
	const root = slideRoot(ctx);
	return root ? readTransition(root) : null;
}

export function setSlideTransition(
	ctx: CommandContext,
	value: Transition | null,
	label: string,
): boolean {
	const root = slideRoot(ctx);
	if (!root) return false;
	return ctx.editor.transact(label, [ctx.slide.partPath], () => writeTransition(root, value), {
		rebuild: false,
	});
}

export function slideAnimations(ctx: CommandContext | null): AnimationEntry[] {
	if (!ctx) return [];
	const root = slideRoot(ctx);
	return root ? readAnimations(root) : [];
}

/** The animations on the shapes currently selected. */
export function selectedAnimations(ctx: CommandContext | null): AnimationEntry[] {
	if (!ctx) return [];
	const ids = new Set(selectedShapes(ctx).map((shape) => shape.id));
	return slideAnimations(ctx).filter((entry) => ids.has(entry.shapeId));
}

export function animateSelection(
	ctx: CommandContext,
	effect: EffectKind,
	trigger: Trigger,
	label: string,
): boolean {
	const root = slideRoot(ctx);
	const shapes = selectedShapes(ctx);
	if (!root || shapes.length === 0) return false;
	return ctx.editor.transact(
		label,
		[ctx.slide.partPath],
		() => {
			let added = false;
			for (const shape of shapes) {
				// Re-animating a shape replaces what it had, which is what PowerPoint
				// does too — two entrances on one shape is never what was meant.
				removeAnimationsFor(root, shape.id);
				if (addAnimation(root, shape.id, effect, trigger)) added = true;
			}
			return added;
		},
		{ rebuild: false },
	);
}

export function removeSelectionAnimation(ctx: CommandContext, label: string): boolean {
	const root = slideRoot(ctx);
	const shapes = selectedShapes(ctx);
	if (!root || shapes.length === 0) return false;
	return ctx.editor.transact(
		label,
		[ctx.slide.partPath],
		() => shapes.reduce((count, shape) => count + removeAnimationsFor(root, shape.id), 0) > 0,
		{ rebuild: false },
	);
}

export function clearSlideAnimations(ctx: CommandContext, label: string): boolean {
	const root = slideRoot(ctx);
	if (!root) return false;
	return ctx.editor.transact(label, [ctx.slide.partPath], () => clearAnimations(root), {
		rebuild: false,
	});
}

/** Move the selected shape's animation one place earlier or later. */
export function moveSelectionAnimation(
	ctx: CommandContext,
	by: -1 | 1,
	label: string,
): boolean {
	const root = slideRoot(ctx);
	if (!root) return false;
	const mine = selectedAnimations(ctx);
	if (mine.length !== 1) return false;
	const from = mine[0].index;
	return ctx.editor.transact(
		label,
		[ctx.slide.partPath],
		() => moveAnimation(root, from, from + by),
		{ rebuild: false },
	);
}

export function canReorderAnimations(ctx: CommandContext | null): boolean {
	if (!ctx) return false;
	const root = slideRoot(ctx);
	return root ? canReorder(root) : false;
}
