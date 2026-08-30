import type { Frame, Shape } from "../pptx/types";
import { shapeRegistry } from "../render/renderSlide";
import { writeShapeFrame } from "./geometryWrite";
import { type ElementHistory, type Snapshot, capture } from "./History";

export interface ShapeEditorOptions {
	/** False while a text box is being edited, so the two never fight over a drag. */
	isEnabled: () => boolean;
	/** Current CSS scale of the slide element, for converting pointer deltas. */
	getScale: () => number;
	history: ElementHistory;
	onChange: (part: string) => void;
}

type HandleId = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";

const HANDLES: { id: HandleId; fx: number; fy: number; cursor: string }[] = [
	{ id: "nw", fx: 0, fy: 0, cursor: "nwse-resize" },
	{ id: "n", fx: 0.5, fy: 0, cursor: "ns-resize" },
	{ id: "ne", fx: 1, fy: 0, cursor: "nesw-resize" },
	{ id: "e", fx: 1, fy: 0.5, cursor: "ew-resize" },
	{ id: "se", fx: 1, fy: 1, cursor: "nwse-resize" },
	{ id: "s", fx: 0.5, fy: 1, cursor: "ns-resize" },
	{ id: "sw", fx: 0, fy: 1, cursor: "nesw-resize" },
	{ id: "w", fx: 0, fy: 0.5, cursor: "ew-resize" },
];

/** Ignore drags under this many screen pixels so a click stays a click. */
const DRAG_THRESHOLD = 2;
const MIN_SIZE = 6;
/** Snap distance, in screen pixels. */
const SNAP = 6;

interface DragState {
	kind: "move" | "resize";
	handle: HandleId | null;
	pointerId: number;
	originX: number;
	originY: number;
	start: Frame;
	snapshot: Snapshot;
	guides: { xs: number[]; ys: number[] };
	moved: boolean;
}

/**
 * Click to select a shape, drag to move it, drag a handle to resize it.
 *
 * Everything happens in slide coordinates: pointer deltas are divided by the
 * view's zoom, so dragging feels identical at any zoom level and the numbers
 * written back to the XML never depend on how the deck was being viewed.
 */
export class ShapeEditor {
	private slideEl: HTMLElement | null = null;
	private selectedEl: HTMLElement | null = null;
	private selectedShape: Shape | null = null;
	private overlay: HTMLElement | null = null;
	private guideLayer: HTMLElement | null = null;
	private drag: DragState | null = null;

	constructor(private readonly options: ShapeEditorOptions) {}

	get hasSelection(): boolean {
		return this.selectedShape !== null;
	}

	attach(slideEl: HTMLElement): void {
		slideEl.addEventListener("pointerdown", this.onPointerDown);
	}

	detach(slideEl: HTMLElement): void {
		slideEl.removeEventListener("pointerdown", this.onPointerDown);
		if (this.slideEl === slideEl) this.deselect();
	}

	deselect(): void {
		this.overlay?.remove();
		this.guideLayer?.remove();
		this.overlay = null;
		this.guideLayer = null;
		this.selectedEl = null;
		this.selectedShape = null;
		this.drag = null;
	}

	/** Re-place the selection overlay after the view's zoom changed. */
	refresh(): void {
		if (this.selectedShape) this.positionOverlay(this.selectedShape.frame);
	}

	// ------------------------------------------------------------ selection

	private onPointerDown = (event: PointerEvent): void => {
		if (!this.options.isEnabled() || event.button !== 0) return;
		const target = event.target;
		if (!(target instanceof HTMLElement)) return;
		const slideEl = event.currentTarget as HTMLElement;

		const handle = target.closest<HTMLElement>("[data-handle]");
		if (handle && this.selectedShape) {
			this.beginDrag(event, "resize", handle.dataset.handle as HandleId);
			return;
		}

		const shapeEl = target.closest<HTMLElement>('[data-selectable="1"]');
		if (!shapeEl || !slideEl.contains(shapeEl)) {
			this.deselect();
			return;
		}
		const shape = shapeRegistry.get(shapeEl);
		if (!shape) {
			this.deselect();
			return;
		}

		if (this.selectedEl !== shapeEl) {
			this.deselect();
			this.slideEl = slideEl;
			this.selectedEl = shapeEl;
			this.selectedShape = shape;
			this.buildOverlay();
		}
		this.beginDrag(event, "move", null);
	};

	private buildOverlay(): void {
		const slideEl = this.slideEl;
		const shape = this.selectedShape;
		if (!slideEl || !shape) return;

		this.guideLayer = slideEl.createDiv({ cls: "pptx-guides" });
		Object.assign(this.guideLayer.style, {
			position: "absolute",
			inset: "0",
			pointerEvents: "none",
			zIndex: "40",
		});

		const overlay = slideEl.createDiv({ cls: "pptx-selection" });
		Object.assign(overlay.style, {
			position: "absolute",
			pointerEvents: "none",
			zIndex: "50",
			boxSizing: "border-box",
		});
		for (const handle of HANDLES) {
			const el = overlay.createDiv({ cls: "pptx-handle" });
			el.dataset.handle = handle.id;
			Object.assign(el.style, {
				position: "absolute",
				pointerEvents: "auto",
				cursor: handle.cursor,
				boxSizing: "border-box",
			});
		}
		this.overlay = overlay;
		this.positionOverlay(shape.frame);
	}

	private positionOverlay(frame: Frame): void {
		const overlay = this.overlay;
		if (!overlay) return;
		const scale = Math.max(this.options.getScale(), 0.05);
		// Handles are sized in slide units so that they come out a constant size
		// on screen whatever the zoom.
		const size = 9 / scale;
		const border = 1.5 / scale;

		Object.assign(overlay.style, {
			left: `${frame.x}px`,
			top: `${frame.y}px`,
			width: `${frame.w}px`,
			height: `${frame.h}px`,
			outline: `${border}px solid var(--interactive-accent, #2f6fed)`,
			transform: frame.rot ? `rotate(${frame.rot}deg)` : "",
			transformOrigin: "center center",
		});

		for (const handle of HANDLES) {
			const el = overlay.querySelector<HTMLElement>(`[data-handle="${handle.id}"]`);
			if (!el) continue;
			Object.assign(el.style, {
				width: `${size}px`,
				height: `${size}px`,
				left: `${frame.w * handle.fx - size / 2}px`,
				top: `${frame.h * handle.fy - size / 2}px`,
				border: `${border}px solid var(--interactive-accent, #2f6fed)`,
				background: "var(--background-primary, #fff)",
				borderRadius: `${size / 4}px`,
			});
		}
	}

	// ---------------------------------------------------------------- drag

	private beginDrag(event: PointerEvent, kind: "move" | "resize", handle: HandleId | null): void {
		const shape = this.selectedShape;
		const source = shape?.source;
		if (!shape || !source) return;
		// Only resize swallows the default action. Calling preventDefault on a move
		// can suppress the follow-up dblclick that opens the text editor.
		if (kind === "resize") event.preventDefault();
		event.stopPropagation();

		this.drag = {
			kind,
			handle,
			pointerId: event.pointerId,
			originX: event.clientX,
			originY: event.clientY,
			start: { ...shape.frame },
			snapshot: capture(source),
			guides: this.collectGuides(shape),
			moved: false,
		};

		window.addEventListener("pointermove", this.onPointerMove);
		window.addEventListener("pointerup", this.onPointerUp);
		window.addEventListener("pointercancel", this.onPointerUp);
	}

	/** Edges and centres of the other shapes, plus the slide's own, to snap against. */
	private collectGuides(exclude: Shape): { xs: number[]; ys: number[] } {
		const slideEl = this.slideEl;
		const xs: number[] = [];
		const ys: number[] = [];
		if (!slideEl) return { xs, ys };

		const width = slideEl.offsetWidth || parseFloat(slideEl.style.width) || 0;
		const height = slideEl.offsetHeight || parseFloat(slideEl.style.height) || 0;
		xs.push(0, width / 2, width);
		ys.push(0, height / 2, height);

		for (const el of Array.from(slideEl.querySelectorAll<HTMLElement>('[data-selectable="1"]'))) {
			const other = shapeRegistry.get(el);
			if (!other || other === exclude || other.hidden || other.frame.rot) continue;
			const f = other.frame;
			xs.push(f.x, f.x + f.w / 2, f.x + f.w);
			ys.push(f.y, f.y + f.h / 2, f.y + f.h);
		}
		return { xs, ys };
	}

	private onPointerMove = (event: PointerEvent): void => {
		const drag = this.drag;
		const shape = this.selectedShape;
		if (!drag || !shape || event.pointerId !== drag.pointerId) return;

		const dxScreen = event.clientX - drag.originX;
		const dyScreen = event.clientY - drag.originY;
		if (!drag.moved && Math.hypot(dxScreen, dyScreen) < DRAG_THRESHOLD) return;
		drag.moved = true;

		const scale = Math.max(this.options.getScale(), 0.05);
		const dx = dxScreen / scale;
		const dy = dyScreen / scale;
		const snapping = !event.altKey && drag.start.rot === 0;

		const frame =
			drag.kind === "move"
				? this.computeMove(drag, dx, dy, snapping, scale)
				: this.computeResize(drag, dx, dy, snapping, scale, event.shiftKey);

		this.applyLiveFrame(frame);
	};

	private computeMove(
		drag: DragState,
		dx: number,
		dy: number,
		snapping: boolean,
		scale: number,
	): Frame {
		const start = drag.start;
		let x = start.x + dx;
		let y = start.y + dy;
		const matched: { xs: number[]; ys: number[] } = { xs: [], ys: [] };

		if (snapping) {
			const threshold = SNAP / scale;
			const xFit = bestSnap([x, x + start.w / 2, x + start.w], drag.guides.xs, threshold);
			if (xFit) {
				x += xFit.delta;
				matched.xs.push(xFit.guide);
			}
			const yFit = bestSnap([y, y + start.h / 2, y + start.h], drag.guides.ys, threshold);
			if (yFit) {
				y += yFit.delta;
				matched.ys.push(yFit.guide);
			}
		}

		this.drawGuides(matched);
		return { ...start, x, y };
	}

	private computeResize(
		drag: DragState,
		dx: number,
		dy: number,
		snapping: boolean,
		scale: number,
		keepAspect: boolean,
	): Frame {
		const start = drag.start;
		const handle = HANDLES.find((h) => h.id === drag.handle);
		if (!handle) return start;

		const movesLeft = handle.fx === 0;
		const movesRight = handle.fx === 1;
		const movesTop = handle.fy === 0;
		const movesBottom = handle.fy === 1;

		// Rotated shapes resize along their own axes, so the pointer delta has to
		// be rotated into the shape's frame first.
		const theta = (start.rot * Math.PI) / 180;
		const cos = Math.cos(theta);
		const sin = Math.sin(theta);
		const localDx = start.rot ? dx * cos + dy * sin : dx;
		const localDy = start.rot ? -dx * sin + dy * cos : dy;

		let w = start.w + (movesRight ? localDx : movesLeft ? -localDx : 0);
		let h = start.h + (movesBottom ? localDy : movesTop ? -localDy : 0);

		if (keepAspect && movesLeft !== movesRight && movesTop !== movesBottom && start.h > 0) {
			const ratio = start.w / start.h;
			if (Math.abs(w - start.w) > Math.abs(h - start.h)) h = w / ratio;
			else w = h * ratio;
		}

		w = Math.max(MIN_SIZE, w);
		h = Math.max(MIN_SIZE, h);

		if (start.rot === 0) {
			let x = movesLeft ? start.x + start.w - w : start.x;
			let y = movesTop ? start.y + start.h - h : start.y;
			const matched: { xs: number[]; ys: number[] } = { xs: [], ys: [] };
			if (snapping) {
				const threshold = SNAP / scale;
				if (movesLeft || movesRight) {
					const edge = movesLeft ? x : x + w;
					const fit = bestSnap([edge], drag.guides.xs, threshold);
					if (fit) {
						matched.xs.push(fit.guide);
						if (movesLeft) {
							x += fit.delta;
							w -= fit.delta;
						} else {
							w += fit.delta;
						}
					}
				}
				if (movesTop || movesBottom) {
					const edge = movesTop ? y : y + h;
					const fit = bestSnap([edge], drag.guides.ys, threshold);
					if (fit) {
						matched.ys.push(fit.guide);
						if (movesTop) {
							y += fit.delta;
							h -= fit.delta;
						} else {
							h += fit.delta;
						}
					}
				}
			}
			this.drawGuides(matched);
			return { ...start, x, y, w: Math.max(MIN_SIZE, w), h: Math.max(MIN_SIZE, h) };
		}

		// Keep the anchor corner still on screen while the box grows around it.
		const anchorX = movesLeft ? 1 : 0;
		const anchorY = movesTop ? 1 : 0;
		const a0x = (anchorX ? start.w : 0) - start.w / 2;
		const a0y = (anchorY ? start.h : 0) - start.h / 2;
		const a1x = (anchorX ? w : 0) - w / 2;
		const a1y = (anchorY ? h : 0) - h / 2;
		const shiftX = (a0x - a1x) * cos - (a0y - a1y) * sin;
		const shiftY = (a0x - a1x) * sin + (a0y - a1y) * cos;
		const cx = start.x + start.w / 2 + shiftX;
		const cy = start.y + start.h / 2 + shiftY;

		this.drawGuides({ xs: [], ys: [] });
		return { ...start, x: cx - w / 2, y: cy - h / 2, w, h };
	}

	/** Update the live DOM without touching the XML. */
	private applyLiveFrame(frame: Frame): void {
		const el = this.selectedEl;
		const shape = this.selectedShape;
		if (!el || !shape) return;
		el.style.left = `${frame.x}px`;
		el.style.top = `${frame.y}px`;
		el.style.width = `${frame.w}px`;
		el.style.height = `${frame.h}px`;

		if (shape.kind === "group") {
			// A group's children live in their own coordinate space; rescale it so
			// they follow the group the way PowerPoint scales them.
			const inner = el.firstElementChild;
			if (inner instanceof HTMLElement) {
				const sx = shape.childOffset.w > 0 ? frame.w / shape.childOffset.w : 1;
				const sy = shape.childOffset.h > 0 ? frame.h / shape.childOffset.h : 1;
				inner.style.transform = `scale(${sx}, ${sy}) translate(${-shape.childOffset.x}px, ${-shape.childOffset.y}px)`;
			}
		}

		this.positionOverlay(frame);
		this.pendingFrame = frame;
	}

	private pendingFrame: Frame | null = null;

	private onPointerUp = (event: PointerEvent): void => {
		const drag = this.drag;
		if (!drag || event.pointerId !== drag.pointerId) return;
		window.removeEventListener("pointermove", this.onPointerMove);
		window.removeEventListener("pointerup", this.onPointerUp);
		window.removeEventListener("pointercancel", this.onPointerUp);
		this.drag = null;
		this.clearGuides();

		const frame = this.pendingFrame;
		this.pendingFrame = null;
		if (!drag.moved || !frame) return;
		this.commit(frame, drag.snapshot, drag.kind === "move" ? "Move shape" : "Resize shape");
	};

	private commit(frame: Frame, snapshot: Snapshot, label: string): void {
		const shape = this.selectedShape;
		if (!shape) return;
		if (
			Math.abs(frame.x - shape.frame.x) < 0.01 &&
			Math.abs(frame.y - shape.frame.y) < 0.01 &&
			Math.abs(frame.w - shape.frame.w) < 0.01 &&
			Math.abs(frame.h - shape.frame.h) < 0.01
		) {
			return;
		}

		const part = writeShapeFrame(shape, frame);
		if (!part) return;
		// Keep the in-memory model in step so the selection survives; re-rendering
		// here would drop it and make dragging feel like it reset.
		shape.frame.x = frame.x;
		shape.frame.y = frame.y;
		shape.frame.w = frame.w;
		shape.frame.h = frame.h;

		this.options.history.record(label, part, snapshot);
		this.options.onChange(part);
	}

	// ------------------------------------------------------------- guides

	private drawGuides(matched: { xs: number[]; ys: number[] }): void {
		const layer = this.guideLayer;
		if (!layer) return;
		layer.empty();
		const scale = Math.max(this.options.getScale(), 0.05);
		const thickness = 1 / scale;
		for (const x of matched.xs) {
			const line = layer.createDiv();
			Object.assign(line.style, {
				position: "absolute",
				left: `${x - thickness / 2}px`,
				top: "0",
				width: `${thickness}px`,
				height: "100%",
				background: "#e8590c",
			});
		}
		for (const y of matched.ys) {
			const line = layer.createDiv();
			Object.assign(line.style, {
				position: "absolute",
				top: `${y - thickness / 2}px`,
				left: "0",
				height: `${thickness}px`,
				width: "100%",
				background: "#e8590c",
			});
		}
	}

	private clearGuides(): void {
		this.guideLayer?.empty();
	}

	// ----------------------------------------------------------- keyboard

	/** Arrow keys nudge the selection. Returns true when the key was consumed. */
	handleKey(event: KeyboardEvent): boolean {
		const shape = this.selectedShape;
		if (!shape?.source) return false;

		if (event.key === "Escape") {
			this.deselect();
			return true;
		}

		const step = event.shiftKey ? 10 : 1;
		let dx = 0;
		let dy = 0;
		if (event.key === "ArrowLeft") dx = -step;
		else if (event.key === "ArrowRight") dx = step;
		else if (event.key === "ArrowUp") dy = -step;
		else if (event.key === "ArrowDown") dy = step;
		else return false;

		const snapshot = capture(shape.source);
		const frame = { ...shape.frame, x: shape.frame.x + dx, y: shape.frame.y + dy };
		this.applyLiveFrame(frame);
		this.pendingFrame = null;
		this.commit(frame, snapshot, "Nudge shape");
		return true;
	}
}

/** The smallest adjustment that lands one of `values` on one of `guides`. */
function bestSnap(
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
