import type { Frame, Shape } from "../pptx/types";
import { shapeRegistry } from "../render/renderSlide";
import type { CommandContext } from "./commands";
import { deleteSelection, selectedShapes } from "./commands";
import type { DeckEditor } from "./DeckEditor";
import { writeFrame } from "./geometryWrite";
import type { PartsPatch } from "./History";
import type { Selection } from "./Selection";

export interface ShapeEditorOptions {
	selection: Selection;
	editor: DeckEditor;
	/** False while a text box is being edited, so the two never fight over a drag. */
	isEnabled: () => boolean;
	/** Current CSS scale of the slide element, for converting pointer deltas. */
	getScale: () => number;
	getContext: () => CommandContext | null;
	onContextMenu: (event: MouseEvent) => void;
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

const DRAG_THRESHOLD = 2;
const MIN_SIZE = 6;
const SNAP = 6;

interface MoveDrag {
	kind: "move";
	pointerId: number;
	originX: number;
	originY: number;
	items: { shape: Shape; el: HTMLElement; start: Frame }[];
	before: PartsPatch;
	guides: { xs: number[]; ys: number[] };
	moved: boolean;
	frames: Map<string, Frame>;
}

interface ResizeDrag {
	kind: "resize";
	handle: HandleId;
	pointerId: number;
	originX: number;
	originY: number;
	shape: Shape;
	el: HTMLElement;
	start: Frame;
	before: PartsPatch;
	guides: { xs: number[]; ys: number[] };
	moved: boolean;
	frame: Frame | null;
}

interface MarqueeDrag {
	kind: "marquee";
	pointerId: number;
	originX: number;
	originY: number;
	additive: boolean;
	baseIds: string[];
	moved: boolean;
}

type Drag = MoveDrag | ResizeDrag | MarqueeDrag;

/**
 * Direct manipulation on the slide: select, multi-select, move, resize.
 *
 * All geometry is computed in slide coordinates — pointer deltas are divided by
 * the view's zoom — so dragging feels the same at any zoom and the numbers
 * written into the XML never depend on how the deck was being viewed.
 */
export class ShapeEditor {
	private slideEl: HTMLElement | null = null;
	private overlay: HTMLElement | null = null;
	private guideLayer: HTMLElement | null = null;
	private marqueeEl: HTMLElement | null = null;
	private drag: Drag | null = null;

	constructor(private readonly options: ShapeEditorOptions) {
		options.selection.onChange(() => this.syncOverlay());
	}

	attach(slideEl: HTMLElement): void {
		slideEl.addEventListener("pointerdown", this.onPointerDown);
		slideEl.addEventListener("contextmenu", this.onContextMenu);
	}

	detach(slideEl: HTMLElement): void {
		slideEl.removeEventListener("pointerdown", this.onPointerDown);
		slideEl.removeEventListener("contextmenu", this.onContextMenu);
		if (this.slideEl === slideEl) {
			this.slideEl = null;
			this.overlay = null;
			this.guideLayer = null;
			this.marqueeEl = null;
		}
	}

	/** Tell the editor which rendered slide is currently on screen. */
	setActive(slideEl: HTMLElement | null): void {
		this.slideEl = slideEl;
		this.overlay = null;
		this.guideLayer = null;
		this.marqueeEl = null;
		this.syncOverlay();
	}

	/** Re-place the overlay after a zoom change or a re-render. */
	refresh(): void {
		this.syncOverlay();
	}

	private get slideIndex(): number {
		return Number(this.slideEl?.dataset.slideIndex ?? 0) - 1;
	}

	// ------------------------------------------------------------- overlay

	private ensureLayers(): void {
		const slideEl = this.slideEl;
		if (!slideEl) return;
		if (this.overlay?.isConnected && this.guideLayer?.isConnected) return;

		this.guideLayer = slideEl.createDiv({ cls: "pptx-guides" });
		Object.assign(this.guideLayer.style, {
			position: "absolute",
			inset: "0",
			pointerEvents: "none",
			zIndex: "40",
		});
		this.overlay = slideEl.createDiv({ cls: "pptx-overlay" });
		Object.assign(this.overlay.style, {
			position: "absolute",
			inset: "0",
			pointerEvents: "none",
			zIndex: "50",
		});
	}

	private elementFor(id: string): HTMLElement | null {
		return this.slideEl?.querySelector<HTMLElement>(`[data-shape-id="${CSS.escape(id)}"]`) ?? null;
	}

	private selectedItems(): { shape: Shape; el: HTMLElement }[] {
		const out: { shape: Shape; el: HTMLElement }[] = [];
		if (!this.slideEl || this.options.selection.slideIndex !== this.slideIndex) return out;
		for (const id of this.options.selection.ids) {
			const el = this.elementFor(id);
			const shape = el ? shapeRegistry.get(el) : undefined;
			if (el && shape) out.push({ shape, el });
		}
		return out;
	}

	private syncOverlay(): void {
		if (!this.slideEl) return;
		this.ensureLayers();
		const overlay = this.overlay;
		if (!overlay) return;
		overlay.empty();

		const items = this.selectedItems();
		if (items.length === 0) return;

		const scale = Math.max(this.options.getScale(), 0.05);
		const border = 1.5 / scale;

		for (const { shape } of items) {
			const box = overlay.createDiv({ cls: "pptx-selection" });
			Object.assign(box.style, {
				position: "absolute",
				boxSizing: "border-box",
				pointerEvents: "none",
				left: `${shape.frame.x}px`,
				top: `${shape.frame.y}px`,
				width: `${shape.frame.w}px`,
				height: `${shape.frame.h}px`,
				outline: `${border}px solid var(--interactive-accent, #2f6fed)`,
				transform: shape.frame.rot ? `rotate(${shape.frame.rot}deg)` : "",
				transformOrigin: "center center",
			});

			// Handles only make sense on a single shape: resizing several at once
			// needs a different model than "drag this shape's corner".
			if (items.length !== 1) continue;
			const size = 9 / scale;
			for (const handle of HANDLES) {
				const el = box.createDiv({ cls: "pptx-handle" });
				el.dataset.handle = handle.id;
				Object.assign(el.style, {
					position: "absolute",
					pointerEvents: "auto",
					cursor: handle.cursor,
					boxSizing: "border-box",
					width: `${size}px`,
					height: `${size}px`,
					left: `${shape.frame.w * handle.fx - size / 2}px`,
					top: `${shape.frame.h * handle.fy - size / 2}px`,
					border: `${border}px solid var(--interactive-accent, #2f6fed)`,
					background: "var(--background-primary, #fff)",
					borderRadius: `${size / 4}px`,
				});
			}
		}
	}

	// -------------------------------------------------------------- input

	private onContextMenu = (event: MouseEvent): void => {
		if (!this.options.isEnabled()) return;
		const target = event.target;
		if (target instanceof HTMLElement) {
			const shapeEl = target.closest<HTMLElement>("[data-shape-id]");
			const id = shapeEl?.dataset.shapeId;
			if (id && !this.options.selection.has(id)) {
				this.options.selection.set(this.slideIndex, [id]);
			}
		}
		event.preventDefault();
		this.options.onContextMenu(event);
	};

	private onPointerDown = (event: PointerEvent): void => {
		if (!this.options.isEnabled() || event.button !== 0) return;
		const target = event.target;
		if (!(target instanceof HTMLElement)) return;
		this.slideEl = event.currentTarget as HTMLElement;

		const handle = target.closest<HTMLElement>("[data-handle]");
		if (handle) {
			this.beginResize(event, handle.dataset.handle as HandleId);
			return;
		}

		const shapeEl = target.closest<HTMLElement>('[data-selectable="1"]');
		const additive = event.shiftKey || event.metaKey || event.ctrlKey;

		if (!shapeEl) {
			this.beginMarquee(event, additive);
			return;
		}
		const id = shapeEl.dataset.shapeId;
		if (!id) return;

		if (additive) {
			this.options.selection.toggle(this.slideIndex, id);
			if (!this.options.selection.has(id)) return;
		} else if (!this.options.selection.has(id)) {
			this.options.selection.set(this.slideIndex, [id]);
		}
		this.beginMove(event);
	};

	private listen(): void {
		window.addEventListener("pointermove", this.onPointerMove);
		window.addEventListener("pointerup", this.onPointerUp);
		window.addEventListener("pointercancel", this.onPointerUp);
	}

	private unlisten(): void {
		window.removeEventListener("pointermove", this.onPointerMove);
		window.removeEventListener("pointerup", this.onPointerUp);
		window.removeEventListener("pointercancel", this.onPointerUp);
	}

	private beginMove(event: PointerEvent): void {
		const items = this.selectedItems().filter((i) => i.shape.source);
		if (items.length === 0) return;
		const part = items[0].shape.sourcePart;
		this.drag = {
			kind: "move",
			pointerId: event.pointerId,
			originX: event.clientX,
			originY: event.clientY,
			items: items.map((i) => ({ ...i, start: { ...i.shape.frame } })),
			before: this.options.editor.capture([part]),
			guides: this.collectGuides(new Set(items.map((i) => i.shape.id))),
			moved: false,
			frames: new Map(),
		};
		// Not preventing the default here keeps the follow-up dblclick, which is
		// what opens the text editor.
		event.stopPropagation();
		this.listen();
	}

	private beginResize(event: PointerEvent, handle: HandleId): void {
		const items = this.selectedItems();
		if (items.length !== 1 || !items[0].shape.source) return;
		const { shape, el } = items[0];
		this.drag = {
			kind: "resize",
			handle,
			pointerId: event.pointerId,
			originX: event.clientX,
			originY: event.clientY,
			shape,
			el,
			start: { ...shape.frame },
			before: this.options.editor.capture([shape.sourcePart]),
			guides: this.collectGuides(new Set([shape.id])),
			moved: false,
			frame: null,
		};
		event.preventDefault();
		event.stopPropagation();
		this.listen();
	}

	private beginMarquee(event: PointerEvent, additive: boolean): void {
		this.drag = {
			kind: "marquee",
			pointerId: event.pointerId,
			originX: event.clientX,
			originY: event.clientY,
			additive,
			baseIds: [...this.options.selection.ids],
			moved: false,
		};
		if (!additive) this.options.selection.clear();
		this.listen();
	}

	private onPointerMove = (event: PointerEvent): void => {
		const drag = this.drag;
		if (!drag || event.pointerId !== drag.pointerId) return;
		const dxScreen = event.clientX - drag.originX;
		const dyScreen = event.clientY - drag.originY;
		if (!drag.moved && Math.hypot(dxScreen, dyScreen) < DRAG_THRESHOLD) return;
		drag.moved = true;

		const scale = Math.max(this.options.getScale(), 0.05);
		if (drag.kind === "marquee") {
			this.updateMarquee(drag, event);
			return;
		}
		if (drag.kind === "move") {
			this.updateMove(drag, dxScreen / scale, dyScreen / scale, !event.altKey, scale);
			return;
		}
		this.updateResize(drag, dxScreen / scale, dyScreen / scale, !event.altKey, scale, event.shiftKey);
	};

	private onPointerUp = (event: PointerEvent): void => {
		const drag = this.drag;
		if (!drag || event.pointerId !== drag.pointerId) return;
		this.unlisten();
		this.drag = null;
		this.guideLayer?.empty();
		this.marqueeEl?.remove();
		this.marqueeEl = null;

		if (!drag.moved) return;
		if (drag.kind === "marquee") return;

		if (drag.kind === "move") {
			let changed = false;
			for (const item of drag.items) {
				const frame = drag.frames.get(item.shape.id);
				if (!frame || !item.shape.source) continue;
				if (Math.abs(frame.x - item.start.x) < 0.01 && Math.abs(frame.y - item.start.y) < 0.01) {
					continue;
				}
				writeFrame(item.shape.source, frame);
				applyToModel(item.shape, frame);
				changed = true;
			}
			if (changed) {
				this.options.editor.recordApplied(
					drag.items.length > 1 ? `Move ${drag.items.length} shapes` : "Move shape",
					drag.before,
				);
			}
			return;
		}

		const frame = drag.frame;
		if (!frame || !drag.shape.source) return;
		if (sameFrame(frame, drag.start)) return;
		writeFrame(drag.shape.source, frame);
		applyToModel(drag.shape, frame);
		this.options.editor.recordApplied("Resize shape", drag.before);
	};

	// --------------------------------------------------------------- move

	private updateMove(
		drag: MoveDrag,
		dx: number,
		dy: number,
		snapping: boolean,
		scale: number,
	): void {
		const matched: { xs: number[]; ys: number[] } = { xs: [], ys: [] };
		let adjustX = 0;
		let adjustY = 0;

		// Snap the selection as a whole so a multi-shape drag keeps its shape.
		if (snapping && drag.items.every((i) => i.start.rot === 0)) {
			const threshold = SNAP / scale;
			const left = Math.min(...drag.items.map((i) => i.start.x)) + dx;
			const right = Math.max(...drag.items.map((i) => i.start.x + i.start.w)) + dx;
			const top = Math.min(...drag.items.map((i) => i.start.y)) + dy;
			const bottom = Math.max(...drag.items.map((i) => i.start.y + i.start.h)) + dy;

			const xFit = bestSnap([left, (left + right) / 2, right], drag.guides.xs, threshold);
			if (xFit) {
				adjustX = xFit.delta;
				matched.xs.push(xFit.guide);
			}
			const yFit = bestSnap([top, (top + bottom) / 2, bottom], drag.guides.ys, threshold);
			if (yFit) {
				adjustY = yFit.delta;
				matched.ys.push(yFit.guide);
			}
		}

		for (const item of drag.items) {
			const frame: Frame = {
				...item.start,
				x: item.start.x + dx + adjustX,
				y: item.start.y + dy + adjustY,
			};
			drag.frames.set(item.shape.id, frame);
			applyLive(item.el, item.shape, frame);
		}
		this.drawGuides(matched);
		this.syncOverlayFrames(drag.frames);
	}

	// ------------------------------------------------------------- resize

	private updateResize(
		drag: ResizeDrag,
		dx: number,
		dy: number,
		snapping: boolean,
		scale: number,
		keepAspect: boolean,
	): void {
		const start = drag.start;
		const handle = HANDLES.find((h) => h.id === drag.handle);
		if (!handle) return;

		const movesLeft = handle.fx === 0;
		const movesRight = handle.fx === 1;
		const movesTop = handle.fy === 0;
		const movesBottom = handle.fy === 1;

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

		let frame: Frame;
		const matched: { xs: number[]; ys: number[] } = { xs: [], ys: [] };

		if (start.rot === 0) {
			let x = movesLeft ? start.x + start.w - w : start.x;
			let y = movesTop ? start.y + start.h - h : start.y;
			if (snapping) {
				const threshold = SNAP / scale;
				if (movesLeft || movesRight) {
					const fit = bestSnap([movesLeft ? x : x + w], drag.guides.xs, threshold);
					if (fit) {
						matched.xs.push(fit.guide);
						if (movesLeft) {
							x += fit.delta;
							w -= fit.delta;
						} else w += fit.delta;
					}
				}
				if (movesTop || movesBottom) {
					const fit = bestSnap([movesTop ? y : y + h], drag.guides.ys, threshold);
					if (fit) {
						matched.ys.push(fit.guide);
						if (movesTop) {
							y += fit.delta;
							h -= fit.delta;
						} else h += fit.delta;
					}
				}
			}
			frame = { ...start, x, y, w: Math.max(MIN_SIZE, w), h: Math.max(MIN_SIZE, h) };
		} else {
			// Hold the anchor corner still on screen while the box grows around it.
			const anchorX = movesLeft ? 1 : 0;
			const anchorY = movesTop ? 1 : 0;
			const a0x = (anchorX ? start.w : 0) - start.w / 2;
			const a0y = (anchorY ? start.h : 0) - start.h / 2;
			const a1x = (anchorX ? w : 0) - w / 2;
			const a1y = (anchorY ? h : 0) - h / 2;
			const cx = start.x + start.w / 2 + (a0x - a1x) * cos - (a0y - a1y) * sin;
			const cy = start.y + start.h / 2 + (a0x - a1x) * sin + (a0y - a1y) * cos;
			frame = { ...start, x: cx - w / 2, y: cy - h / 2, w, h };
		}

		drag.frame = frame;
		applyLive(drag.el, drag.shape, frame);
		this.drawGuides(matched);
		this.syncOverlayFrames(new Map([[drag.shape.id, frame]]));
	}

	// ------------------------------------------------------------ marquee

	private updateMarquee(drag: MarqueeDrag, event: PointerEvent): void {
		const slideEl = this.slideEl;
		if (!slideEl) return;
		const rect = slideEl.getBoundingClientRect();
		const scale = Math.max(this.options.getScale(), 0.05);
		const toSlide = (cx: number, cy: number) => ({
			x: (cx - rect.left) / scale,
			y: (cy - rect.top) / scale,
		});
		const a = toSlide(drag.originX, drag.originY);
		const b = toSlide(event.clientX, event.clientY);
		const box = {
			x: Math.min(a.x, b.x),
			y: Math.min(a.y, b.y),
			w: Math.abs(a.x - b.x),
			h: Math.abs(a.y - b.y),
		};

		this.ensureLayers();
		if (!this.marqueeEl && this.guideLayer) {
			this.marqueeEl = this.guideLayer.createDiv({ cls: "pptx-marquee" });
			Object.assign(this.marqueeEl.style, {
				position: "absolute",
				border: `${1 / scale}px solid var(--interactive-accent, #2f6fed)`,
				background: "rgba(47, 111, 237, 0.12)",
			});
		}
		if (this.marqueeEl) {
			Object.assign(this.marqueeEl.style, {
				left: `${box.x}px`,
				top: `${box.y}px`,
				width: `${box.w}px`,
				height: `${box.h}px`,
			});
		}

		const hits = new Set(drag.additive ? drag.baseIds : []);
		for (const el of Array.from(slideEl.querySelectorAll<HTMLElement>("[data-shape-id]"))) {
			const shape = shapeRegistry.get(el);
			if (!shape || shape.hidden) continue;
			if (intersects(shape.frame, box)) hits.add(shape.id);
		}
		this.options.selection.set(this.slideIndex, hits);
	}

	// ------------------------------------------------------------- guides

	private collectGuides(exclude: Set<string>): { xs: number[]; ys: number[] } {
		const slideEl = this.slideEl;
		const xs: number[] = [];
		const ys: number[] = [];
		if (!slideEl) return { xs, ys };
		const width = parseFloat(slideEl.style.width) || slideEl.offsetWidth;
		const height = parseFloat(slideEl.style.height) || slideEl.offsetHeight;
		xs.push(0, width / 2, width);
		ys.push(0, height / 2, height);
		for (const el of Array.from(slideEl.querySelectorAll<HTMLElement>("[data-shape-id]"))) {
			const shape = shapeRegistry.get(el);
			if (!shape || exclude.has(shape.id) || shape.hidden || shape.frame.rot) continue;
			const f = shape.frame;
			xs.push(f.x, f.x + f.w / 2, f.x + f.w);
			ys.push(f.y, f.y + f.h / 2, f.y + f.h);
		}
		return { xs, ys };
	}

	private drawGuides(matched: { xs: number[]; ys: number[] }): void {
		const layer = this.guideLayer;
		if (!layer) return;
		for (const el of Array.from(layer.querySelectorAll(".pptx-guide"))) el.remove();
		const scale = Math.max(this.options.getScale(), 0.05);
		const thickness = 1 / scale;
		for (const x of matched.xs) {
			const line = layer.createDiv({ cls: "pptx-guide" });
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
			const line = layer.createDiv({ cls: "pptx-guide" });
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

	/** Move the outlines with the shapes during a drag, without a full rebuild. */
	private syncOverlayFrames(frames: Map<string, Frame>): void {
		const overlay = this.overlay;
		if (!overlay) return;
		const boxes = Array.from(overlay.querySelectorAll<HTMLElement>(".pptx-selection"));
		const items = this.selectedItems();
		items.forEach((item, i) => {
			const frame = frames.get(item.shape.id);
			const box = boxes[i];
			if (!frame || !box) return;
			box.style.left = `${frame.x}px`;
			box.style.top = `${frame.y}px`;
			box.style.width = `${frame.w}px`;
			box.style.height = `${frame.h}px`;
			const scale = Math.max(this.options.getScale(), 0.05);
			const size = 9 / scale;
			for (const handle of HANDLES) {
				const el = box.querySelector<HTMLElement>(`[data-handle="${handle.id}"]`);
				if (!el) continue;
				el.style.left = `${frame.w * handle.fx - size / 2}px`;
				el.style.top = `${frame.h * handle.fy - size / 2}px`;
			}
		});
	}

	// ----------------------------------------------------------- keyboard

	/** Returns true when the key was consumed by the selection. */
	handleKey(event: KeyboardEvent): boolean {
		if (!this.options.isEnabled()) return false;
		const ctx = this.options.getContext();

		if (event.key === "Escape" && !this.options.selection.isEmpty) {
			this.options.selection.clear();
			return true;
		}
		if ((event.key === "Delete" || event.key === "Backspace") && ctx) {
			if (this.options.selection.isEmpty) return false;
			deleteSelection(ctx);
			return true;
		}
		if (event.metaKey || event.ctrlKey || event.altKey) return false;

		const step = event.shiftKey ? 10 : 1;
		let dx = 0;
		let dy = 0;
		if (event.key === "ArrowLeft") dx = -step;
		else if (event.key === "ArrowRight") dx = step;
		else if (event.key === "ArrowUp") dy = -step;
		else if (event.key === "ArrowDown") dy = step;
		else return false;

		const items = this.selectedItems().filter((i) => i.shape.source);
		if (items.length === 0) return false;

		const before = this.options.editor.capture([items[0].shape.sourcePart]);
		for (const item of items) {
			const frame = { ...item.shape.frame, x: item.shape.frame.x + dx, y: item.shape.frame.y + dy };
			writeFrame(item.shape.source!, frame);
			applyToModel(item.shape, frame);
			applyLive(item.el, item.shape, frame);
		}
		this.options.editor.recordApplied("Nudge", before);
		this.syncOverlay();
		return true;
	}

	/** Select everything on the current slide. */
	selectAll(): void {
		const slideEl = this.slideEl;
		if (!slideEl) return;
		const ids = Array.from(slideEl.querySelectorAll<HTMLElement>("[data-shape-id]"))
			.map((el) => el.dataset.shapeId)
			.filter((id): id is string => Boolean(id));
		this.options.selection.set(this.slideIndex, ids);
	}
}

// ------------------------------------------------------------- helpers

function applyLive(el: HTMLElement, shape: Shape, frame: Frame): void {
	el.style.left = `${frame.x}px`;
	el.style.top = `${frame.y}px`;
	el.style.width = `${frame.w}px`;
	el.style.height = `${frame.h}px`;
	if (shape.kind === "group") {
		const inner = el.firstElementChild;
		if (inner instanceof HTMLElement) {
			const sx = shape.childOffset.w > 0 ? frame.w / shape.childOffset.w : 1;
			const sy = shape.childOffset.h > 0 ? frame.h / shape.childOffset.h : 1;
			inner.style.transform = `scale(${sx}, ${sy}) translate(${-shape.childOffset.x}px, ${-shape.childOffset.y}px)`;
		}
	}
}

/** Keep the in-memory model in step so the selection survives without a rebuild. */
function applyToModel(shape: Shape, frame: Frame): void {
	shape.frame.x = frame.x;
	shape.frame.y = frame.y;
	shape.frame.w = frame.w;
	shape.frame.h = frame.h;
}

function sameFrame(a: Frame, b: Frame): boolean {
	return (
		Math.abs(a.x - b.x) < 0.01 &&
		Math.abs(a.y - b.y) < 0.01 &&
		Math.abs(a.w - b.w) < 0.01 &&
		Math.abs(a.h - b.h) < 0.01
	);
}

function intersects(a: Frame, b: { x: number; y: number; w: number; h: number }): boolean {
	return a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;
}

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

/** Exported for the ribbon, which needs to know whether commands apply. */
export function selectionSummary(ctx: CommandContext | null): {
	count: number;
	hasGroup: boolean;
} {
	if (!ctx) return { count: 0, hasGroup: false };
	const shapes = selectedShapes(ctx);
	return { count: shapes.length, hasGroup: shapes.some((s) => s.kind === "group") };
}
