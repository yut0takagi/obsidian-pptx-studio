/**
 * The geometry behind direct manipulation, with nothing in it that touches the
 * DOM or the deck.
 *
 * Resizing a rotated shape and snapping a drag are the two places where an
 * arithmetic slip does not throw, does not fail to save, and does not look
 * wrong until someone notices a box has crept. Keeping them here means they can
 * be tested against numbers rather than against a gesture.
 */
import type { Frame, Shape } from "../pptx/types";

export type HandleId = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";

export interface Handle {
	id: HandleId;
	/** Where the handle sits on the box, as a fraction of its width and height. */
	fx: number;
	fy: number;
	cursor: string;
}

export const HANDLES: Handle[] = [
	{ id: "nw", fx: 0, fy: 0, cursor: "nwse-resize" },
	{ id: "n", fx: 0.5, fy: 0, cursor: "ns-resize" },
	{ id: "ne", fx: 1, fy: 0, cursor: "nesw-resize" },
	{ id: "e", fx: 1, fy: 0.5, cursor: "ew-resize" },
	{ id: "se", fx: 1, fy: 1, cursor: "nwse-resize" },
	{ id: "s", fx: 0.5, fy: 1, cursor: "ns-resize" },
	{ id: "sw", fx: 0, fy: 1, cursor: "nesw-resize" },
	{ id: "w", fx: 0, fy: 0.5, cursor: "ew-resize" },
];

export const DRAG_THRESHOLD = 2;
export const MIN_SIZE = 6;
export const SNAP = 6;
/** Rotation snaps to this many degrees while Shift is held. */
export const ROTATE_SNAP = 15;

/** Resize a rotated shape along its own axes, holding the anchor corner still. */
export function rotatedResize(
	start: Frame,
	handle: { fx: number; fy: number },
	dx: number,
	dy: number,
	keepAspect: boolean,
): Frame {
	const theta = (start.rot * Math.PI) / 180;
	const cos = Math.cos(theta);
	const sin = Math.sin(theta);
	const localDx = dx * cos + dy * sin;
	const localDy = -dx * sin + dy * cos;

	const movesLeft = handle.fx === 0;
	const movesRight = handle.fx === 1;
	const movesTop = handle.fy === 0;
	const movesBottom = handle.fy === 1;

	let w = start.w + (movesRight ? localDx : movesLeft ? -localDx : 0);
	let h = start.h + (movesBottom ? localDy : movesTop ? -localDy : 0);
	if (keepAspect && movesLeft !== movesRight && movesTop !== movesBottom && start.h > 0) {
		const ratio = start.w / start.h;
		if (Math.abs(w - start.w) > Math.abs(h - start.h)) h = w / ratio;
		else w = h * ratio;
	}
	w = Math.max(MIN_SIZE, w);
	h = Math.max(MIN_SIZE, h);

	const anchorX = movesLeft ? 1 : 0;
	const anchorY = movesTop ? 1 : 0;
	const a0x = (anchorX ? start.w : 0) - start.w / 2;
	const a0y = (anchorY ? start.h : 0) - start.h / 2;
	const a1x = (anchorX ? w : 0) - w / 2;
	const a1y = (anchorY ? h : 0) - h / 2;
	const cx = start.x + start.w / 2 + (a0x - a1x) * cos - (a0y - a1y) * sin;
	const cy = start.y + start.h / 2 + (a0x - a1x) * sin + (a0y - a1y) * cos;
	return { ...start, x: cx - w / 2, y: cy - h / 2, w, h };
}

/** Keep the in-memory model in step so the selection survives without a rebuild. */
export function applyToModel(shape: Shape, frame: Frame): void {
	shape.frame.x = frame.x;
	shape.frame.y = frame.y;
	shape.frame.w = frame.w;
	shape.frame.h = frame.h;
}

export function unionFrame(frames: Frame[]): Frame {
	const left = Math.min(...frames.map((f) => f.x));
	const top = Math.min(...frames.map((f) => f.y));
	const right = Math.max(...frames.map((f) => f.x + f.w));
	const bottom = Math.max(...frames.map((f) => f.y + f.h));
	return { x: left, y: top, w: right - left, h: bottom - top, rot: 0, flipH: false, flipV: false };
}

export function sameFrame(a: Frame, b: Frame): boolean {
	return (
		Math.abs(a.x - b.x) < 0.01 &&
		Math.abs(a.y - b.y) < 0.01 &&
		Math.abs(a.w - b.w) < 0.01 &&
		Math.abs(a.h - b.h) < 0.01
	);
}

export function intersects(a: Frame, b: { x: number; y: number; w: number; h: number }): boolean {
	return a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;
}

export function bestSnap(
	values: number[],
	guides: number[],
	threshold: number,
): { delta: number; guide: number } | null {
	let best: { delta: number; guide: number } | null = null;
	for (const value of values) {
		for (const guide of guides) {
			const delta = guide - value;
			if (Math.abs(delta) > threshold) continue;
			if (!best || Math.abs(delta) < Math.abs(best.delta)) best = { delta, guide };
		}
	}
	return best;
}
