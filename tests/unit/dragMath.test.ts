/**
 * The arithmetic behind moving, resizing and snapping.
 *
 * None of this is reached by the smoke test, which drives the editor through
 * commands rather than through gestures — so until now the only check on a
 * resize was whether it looked right by hand. The invariant worth holding is
 * the one a user would name: the corner you are not dragging does not move.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";

import {
	MIN_SIZE,
	bestSnap,
	intersects,
	rotatedResize,
	sameFrame,
	unionFrame,
} from "../../src/edit/dragMath";
import type { Frame } from "../../src/pptx/types";

function frame(x: number, y: number, w: number, h: number, rot = 0): Frame {
	return { x, y, w, h, rot, flipH: false, flipV: false };
}

/**
 * Where a corner of a frame actually sits, derived independently of the code
 * under test: rotate the corner's offset from the centre and add it back.
 */
function cornerAt(f: Frame, fx: number, fy: number): { x: number; y: number } {
	const theta = (f.rot * Math.PI) / 180;
	const cos = Math.cos(theta);
	const sin = Math.sin(theta);
	const lx = fx * f.w - f.w / 2;
	const ly = fy * f.h - f.h / 2;
	return {
		x: f.x + f.w / 2 + lx * cos - ly * sin,
		y: f.y + f.h / 2 + lx * sin + ly * cos,
	};
}

function assertClose(actual: number, expected: number, message: string): void {
	assert.ok(
		Math.abs(actual - expected) < 1e-9,
		`${message}: expected ${expected}, got ${actual}`,
	);
}

const SE = { fx: 1, fy: 1 };
const NW = { fx: 0, fy: 0 };
const N = { fx: 0.5, fy: 0 };

describe("rotatedResize", () => {
	it("grows an unrotated box from its bottom-right, leaving the origin alone", () => {
		const out = rotatedResize(frame(100, 50, 200, 100), SE, 40, 20, false);
		assert.deepEqual(
			{ x: out.x, y: out.y, w: out.w, h: out.h },
			{ x: 100, y: 50, w: 240, h: 120 },
		);
	});

	it("moves the origin when the top-left is dragged, holding the far corner", () => {
		const start = frame(100, 50, 200, 100);
		const out = rotatedResize(start, NW, 40, 20, false);
		assert.deepEqual({ w: out.w, h: out.h }, { w: 160, h: 80 });
		assertClose(out.x + out.w, start.x + start.w, "right edge");
		assertClose(out.y + out.h, start.y + start.h, "bottom edge");
	});

	it("only moves the axis its handle owns", () => {
		const out = rotatedResize(frame(100, 50, 200, 100), N, 40, 20, false);
		assert.equal(out.w, 200);
		assert.equal(out.h, 80);
	});

	it("never shrinks below the minimum, however far the pointer goes", () => {
		const out = rotatedResize(frame(100, 50, 200, 100), SE, -1000, -1000, false);
		assert.equal(out.w, MIN_SIZE);
		assert.equal(out.h, MIN_SIZE);
	});

	it("keeps the aspect ratio from the axis that moved further", () => {
		const out = rotatedResize(frame(100, 50, 200, 100), SE, 100, 0, true);
		assert.equal(out.w, 300);
		assert.equal(out.h, 150);
	});

	it("ignores the aspect ratio on an edge handle, which has only one axis", () => {
		const out = rotatedResize(frame(100, 50, 200, 100), N, 0, 20, true);
		assert.equal(out.w, 200);
		assert.equal(out.h, 80);
	});

	it("holds the anchor corner of a rotated shape exactly where it was", () => {
		for (const [handle, anchor] of [
			[SE, { fx: 0, fy: 0 }],
			[NW, { fx: 1, fy: 1 }],
		] as const) {
			for (const rot of [30, 90, 200, -45]) {
				const start = frame(100, 50, 200, 100, rot);
				const out = rotatedResize(start, handle, 37, -19, false);
				const was = cornerAt(start, anchor.fx, anchor.fy);
				const now = cornerAt(out, anchor.fx, anchor.fy);
				assertClose(now.x, was.x, `anchor x at ${rot}deg`);
				assertClose(now.y, was.y, `anchor y at ${rot}deg`);
			}
		}
	});

	it("resizes along the shape's own axes, not the screen's", () => {
		// At 90 degrees the shape's local x runs down the screen, so a downward
		// drag is what widens it.
		const out = rotatedResize(frame(100, 50, 200, 100, 90), SE, 0, 40, false);
		assertClose(out.w, 240, "width");
		assertClose(out.h, 100, "height");
	});

	it("carries the rotation and flips through untouched", () => {
		const start = { ...frame(0, 0, 100, 100, 45), flipH: true };
		const out = rotatedResize(start, SE, 10, 10, false);
		assert.equal(out.rot, 45);
		assert.equal(out.flipH, true);
	});
});

describe("unionFrame", () => {
	it("bounds every frame given", () => {
		const out = unionFrame([frame(10, 20, 30, 40), frame(100, 0, 10, 10)]);
		assert.deepEqual({ x: out.x, y: out.y, w: out.w, h: out.h }, { x: 10, y: 0, w: 100, h: 60 });
	});

	it("is upright, because a bounding box has no rotation of its own", () => {
		const out = unionFrame([frame(0, 0, 10, 10, 45)]);
		assert.equal(out.rot, 0);
		assert.equal(out.flipH, false);
		assert.equal(out.flipV, false);
	});
});

describe("sameFrame", () => {
	it("treats a sub-hundredth difference as no change", () => {
		assert.equal(sameFrame(frame(0, 0, 10, 10), frame(0.005, 0, 10, 10)), true);
	});

	it("notices anything larger", () => {
		assert.equal(sameFrame(frame(0, 0, 10, 10), frame(0.02, 0, 10, 10)), false);
	});

	it("compares position and size only — rotation is written by its own path", () => {
		assert.equal(sameFrame(frame(0, 0, 10, 10, 0), frame(0, 0, 10, 10, 90)), true);
	});
});

describe("intersects", () => {
	it("catches an overlap and a containment", () => {
		assert.equal(intersects(frame(0, 0, 10, 10), { x: 5, y: 5, w: 10, h: 10 }), true);
		assert.equal(intersects(frame(2, 2, 2, 2), { x: 0, y: 0, w: 10, h: 10 }), true);
	});

	it("does not count a shared edge, so a marquee drawn up to a shape misses it", () => {
		assert.equal(intersects(frame(0, 0, 10, 10), { x: 10, y: 0, w: 5, h: 5 }), false);
	});

	it("rejects a box that is clear of it", () => {
		assert.equal(intersects(frame(0, 0, 10, 10), { x: 20, y: 20, w: 5, h: 5 }), false);
	});
});

describe("bestSnap", () => {
	it("takes the closest guide within the threshold", () => {
		assert.deepEqual(bestSnap([100], [104, 98, 130], 6), { delta: -2, guide: 98 });
	});

	it("considers every value, not just the first", () => {
		assert.deepEqual(bestSnap([100, 200], [201], 6), { delta: 1, guide: 201 });
	});

	it("returns null when nothing is near enough, so the drag stays free", () => {
		assert.equal(bestSnap([100], [130], 6), null);
		assert.equal(bestSnap([100], [], 6), null);
		assert.equal(bestSnap([], [100], 6), null);
	});

	it("snaps to a guide exactly on the value without nudging it", () => {
		assert.deepEqual(bestSnap([100], [100], 6), { delta: 0, guide: 100 });
	});
});
