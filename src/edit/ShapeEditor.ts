import type { Frame, Shape } from "../pptx/types";
import { shapeRegistry } from "../render/renderSlide";
import type { CommandContext } from "./commands";
import { deleteSelection, duplicateSelection, selectedShapes } from "./commands";
import type { DeckEditor } from "./DeckEditor";
import { writeFrame, writeRotation } from "./geometryWrite";
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
	/** Enter or F2 on a shape that holds text. */
	onEditText?: (shapeId: string) => void;
	/**
	 * A printable key pressed on a selected shape, which retypes its text.
	 * Returns false when the shape holds no editable text, so the key can fall
	 * through to whatever else wanted it.
	 */
	onTypeText?: (shapeId: string, text: string) => boolean;
	/** Leave text editing, because this press was aimed at the shape itself. */
	onLeaveText?: () => void;
	/** Called when a click lands on a table cell, so the table editor can follow. */
	onCellPointerDown?: (shape: Shape, row: number, column: number, additive: boolean) => void;
	/** User guides to snap against, alongside the other shapes and the slide. */
	extraGuides?: () => { xs: number[]; ys: number[] };
	/** Gives the guide controller first refusal on a press, so guides can be dragged. */
	claimPointer?: (event: PointerEvent) => boolean;
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
/** Rotation snaps to this many degrees while Shift is held. */
const ROTATE_SNAP = 15;

interface DragItem {
	shape: Shape;
	el: HTMLElement;
	start: Frame;
}

interface DragBase {
	pointerId: number;
	originX: number;
	originY: number;
	items: DragItem[];
	before: PartsPatch;
	moved: boolean;
	frames: Map<string, Frame>;
}

interface MoveDrag extends DragBase {
	kind: "move";
	guides: { xs: number[]; ys: number[] };
	/** Alt-drag leaves the original behind and moves a copy. */
	duplicated: boolean;
}

interface ResizeDrag extends DragBase {
	kind: "resize";
	handle: HandleId;
	/** Union of the starting frames: what the handles actually resize. */
	bounds: Frame;
	guides: { xs: number[]; ys: number[] };
}

interface RotateDrag extends DragBase {
	kind: "rotate";
	centerX: number;
	centerY: number;
	startAngle: number;
	rotations: Map<string, number>;
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

type Drag = MoveDrag | ResizeDrag | RotateDrag | MarqueeDrag;

/**
 * Direct manipulation on the slide: select, multi-select, move, resize, rotate.
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

	/**
	 * The empty space around the slide.
	 *
	 * A lasso meant to take in the shapes along the slide's edge naturally
	 * starts outside it, which used to do nothing at all: the only place a
	 * marquee could begin was on the slide's own background.
	 */
	attachBackdrop(el: HTMLElement): void {
		el.addEventListener("pointerdown", this.onBackdropPointerDown);
	}

	detachBackdrop(el: HTMLElement): void {
		el.removeEventListener("pointerdown", this.onBackdropPointerDown);
	}

	private onBackdropPointerDown = (event: PointerEvent): void => {
		if (event.button !== 0 || !this.slideEl) return;
		const target = event.target;
		// Only the backdrop itself: a press that reached a ruler, the toolbar or
		// the slide belongs to whatever it landed on.
		if (
			target !== event.currentTarget &&
			!(isHtmlElement(target) && target.classList.contains("pptx-canvas"))
		) {
			return;
		}
		if (this.editing) this.options.onLeaveText?.();
		else if (!this.options.isEnabled()) return;
		this.beginMarquee(event, event.shiftKey || event.metaKey || event.ctrlKey);
	};

	detach(slideEl: HTMLElement): void {
		slideEl.removeEventListener("pointerdown", this.onPointerDown);
		slideEl.removeEventListener("contextmenu", this.onContextMenu);
		if (this.slideEl === slideEl) this.clearLayers();
	}

	/** Tell the editor which rendered slide is currently on screen. */
	setActive(slideEl: HTMLElement | null): void {
		this.clearLayers();
		this.slideEl = slideEl;
		this.syncOverlay();
	}

	/** True while a text box is open, which the selection is drawn to reflect. */
	private editing = false;

	/**
	 * The text editor opened or closed. A box is either being held — moved,
	 * resized, restyled as a whole — or being typed into, and the border says
	 * which: solid for held, dashed for typing, as in PowerPoint.
	 */
	setEditing(editing: boolean): void {
		if (this.editing === editing) return;
		this.editing = editing;
		this.syncOverlay();
	}

	refresh(): void {
		this.syncOverlay();
	}

	private clearLayers(): void {
		this.overlay?.remove();
		this.guideLayer?.remove();
		this.overlay = null;
		this.guideLayer = null;
		this.marqueeEl = null;
	}

	private get slideIndex(): number {
		return Number(this.slideEl?.dataset.slideIndex ?? 0) - 1;
	}

	// ------------------------------------------------------------- overlay

	private ensureLayers(): void {
		const slideEl = this.slideEl;
		if (!slideEl) return;
		if (this.guideLayer?.isConnected && this.overlay?.isConnected) return;
		this.clearLayers();

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
		if (!this.slideEl) return null;
		for (const el of Array.from(this.slideEl.querySelectorAll<HTMLElement>("[data-shape-id]"))) {
			if (el.dataset.shapeId === id) return el;
		}
		return null;
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

		// Each shape gets a thin outline; the handles live on one box around the
		// whole selection, which is what a multi-shape resize actually scales.
		if (items.length > 1) {
			for (const { shape } of items) {
				const outline = overlay.createDiv({ cls: "pptx-selection is-member" });
				Object.assign(outline.style, {
					position: "absolute",
					pointerEvents: "none",
					boxSizing: "border-box",
					left: `${shape.frame.x}px`,
					top: `${shape.frame.y}px`,
					width: `${shape.frame.w}px`,
					height: `${shape.frame.h}px`,
					outline: `${border}px dashed var(--interactive-accent, #2f6fed)`,
					transform: shape.frame.rot ? `rotate(${shape.frame.rot}deg)` : "",
					transformOrigin: "center center",
				});
			}
		}

		const single = items.length === 1 ? items[0].shape : null;
		const bounds = single ? single.frame : unionFrame(items.map((i) => i.shape.frame));
		const box = overlay.createDiv({ cls: "pptx-selection" });
		Object.assign(box.style, {
			position: "absolute",
			pointerEvents: "none",
			boxSizing: "border-box",
			left: `${bounds.x}px`,
			top: `${bounds.y}px`,
			width: `${bounds.w}px`,
			height: `${bounds.h}px`,
			outline: `${border}px ${this.editing ? "dashed" : "solid"} var(--interactive-accent, #2f6fed)`,
			transform: single?.frame.rot ? `rotate(${single.frame.rot}deg)` : "",
			transformOrigin: "center center",
		});

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
				left: `${bounds.w * handle.fx - size / 2}px`,
				top: `${bounds.h * handle.fy - size / 2}px`,
				border: `${border}px solid var(--interactive-accent, #2f6fed)`,
				background: "var(--background-primary, #fff)",
				borderRadius: `${size / 4}px`,
			});
		}

		// Rotation is only offered for a single shape: rotating a multi-selection
		// about a shared centre is a different operation, and guessing wrong here
		// would scatter the shapes.
		if (single) {
			const stem = box.createDiv({ cls: "pptx-rotate-stem" });
			Object.assign(stem.style, {
				position: "absolute",
				left: `${bounds.w / 2 - border / 2}px`,
				top: `${-size * 2}px`,
				width: `${border}px`,
				height: `${size * 2}px`,
				background: "var(--interactive-accent, #2f6fed)",
				pointerEvents: "none",
			});
			const knob = box.createDiv({ cls: "pptx-rotate" });
			knob.dataset.rotate = "1";
			Object.assign(knob.style, {
				position: "absolute",
				pointerEvents: "auto",
				cursor: "grab",
				boxSizing: "border-box",
				width: `${size}px`,
				height: `${size}px`,
				left: `${bounds.w / 2 - size / 2}px`,
				top: `${-size * 2.8}px`,
				border: `${border}px solid var(--interactive-accent, #2f6fed)`,
				background: "var(--background-primary, #fff)",
				borderRadius: "50%",
			});
		}
	}

	// -------------------------------------------------------------- input

	private onContextMenu = (event: MouseEvent): void => {
		if (!this.options.isEnabled()) return;
		const target = event.target;
		if (isHtmlElement(target)) {
			const shapeEl = target.closest<HTMLElement>("[data-shape-id]");
			const id = shapeEl?.dataset.shapeId;
			if (id && !this.options.selection.has(id)) {
				this.options.selection.set(this.slideIndex, [id]);
			}
		}
		event.preventDefault();
		this.options.onContextMenu(event);
	};

	/**
	 * A press that belongs to the text rather than to the box holding it: inside
	 * the box being edited, and away from its border. The band is measured in
	 * screen pixels because it is a grip for the pointer, not part of the slide.
	 */
	private isTextPress(event: PointerEvent, target: HTMLElement): boolean {
		const box = target.closest<HTMLElement>('[data-editable="1"].is-editing');
		if (!box) return false;
		const rect = box.getBoundingClientRect();
		const edge = 6;
		return !(
			event.clientX - rect.left < edge ||
			rect.right - event.clientX < edge ||
			event.clientY - rect.top < edge ||
			rect.bottom - event.clientY < edge
		);
	}

	private onPointerDown = (event: PointerEvent): void => {
		if (event.button !== 0) return;
		const target = event.target;
		if (!isHtmlElement(target)) return;
		this.slideEl = event.currentTarget as HTMLElement;

		// What was pressed is read before anything else happens: leaving a text
		// edit re-renders the slide, and these elements do not survive that. Ids
		// and screen coordinates do, and they are all the rest of this needs.
		const onRotate = Boolean(target.closest("[data-rotate]"));
		const handleId = target.closest<HTMLElement>("[data-handle]")?.dataset.handle as
			| HandleId
			| undefined;
		const shapeId = target.closest<HTMLElement>('[data-selectable="1"]')?.dataset.shapeId;
		const cellEl = target.closest<HTMLElement>("[data-cell-row]");

		if (this.editing) {
			// While the caret is in a box there are two gestures on it: the text
			// takes presses in its middle, and the box itself takes everything else
			// — its border, its handles, and any other shape on the slide. Leaving
			// the text first is what turns the press back into a plain grab.
			if (!onRotate && !handleId && this.isTextPress(event, target)) return;
			event.preventDefault();
			this.options.onLeaveText?.();
		} else if (!this.options.isEnabled()) {
			return;
		}

		// A press on a guide moves the guide, not whatever is behind it.
		if (!onRotate && !handleId && this.options.claimPointer?.(event)) return;

		if (onRotate) {
			this.beginRotate(event);
			return;
		}
		if (handleId) {
			this.beginResize(event, handleId);
			return;
		}

		const additive = event.shiftKey || event.metaKey || event.ctrlKey;
		if (!shapeId) {
			this.beginMarquee(event, additive);
			return;
		}

		if (additive) {
			this.options.selection.toggle(this.slideIndex, shapeId);
			if (!this.options.selection.has(shapeId)) return;
		} else if (!this.options.selection.has(shapeId)) {
			this.options.selection.set(this.slideIndex, [shapeId]);
		}

		// A click inside a table tells the table editor which cell was hit.
		const shapeEl = this.elementFor(shapeId);
		const shape = shapeEl ? shapeRegistry.get(shapeEl) : undefined;
		if (shape?.kind === "table" && cellEl && this.options.onCellPointerDown) {
			this.options.onCellPointerDown(
				shape,
				Number(cellEl.dataset.cellRow),
				Number(cellEl.dataset.cellCol),
				event.shiftKey,
			);
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

	private dragItems(): DragItem[] {
		return this.selectedItems()
			.filter((i) => i.shape.source)
			.map((i) => ({ ...i, start: { ...i.shape.frame } }));
	}

	private beginMove(event: PointerEvent): void {
		let duplicated = false;
		// Alt-drag copies, the way it does everywhere else: the original stays put
		// and the drag carries the copy.
		if (event.altKey) {
			const ctx = this.options.getContext();
			if (ctx && duplicateSelection(ctx)) duplicated = true;
		}

		const items = this.dragItems();
		if (items.length === 0) return;
		this.drag = {
			kind: "move",
			pointerId: event.pointerId,
			originX: event.clientX,
			originY: event.clientY,
			items,
			before: this.options.editor.capture([items[0].shape.sourcePart]),
			guides: this.collectGuides(new Set(items.map((i) => i.shape.id))),
			moved: duplicated,
			frames: new Map(),
			duplicated,
		};
		event.stopPropagation();
		this.listen();
	}

	private beginResize(event: PointerEvent, handle: HandleId): void {
		const items = this.dragItems();
		if (items.length === 0) return;
		this.drag = {
			kind: "resize",
			handle,
			pointerId: event.pointerId,
			originX: event.clientX,
			originY: event.clientY,
			items,
			bounds: items.length === 1 ? { ...items[0].start } : unionFrame(items.map((i) => i.start)),
			before: this.options.editor.capture([items[0].shape.sourcePart]),
			guides: this.collectGuides(new Set(items.map((i) => i.shape.id))),
			moved: false,
			frames: new Map(),
		};
		event.preventDefault();
		event.stopPropagation();
		this.listen();
	}

	private beginRotate(event: PointerEvent): void {
		const items = this.dragItems();
		if (items.length !== 1 || !this.slideEl) return;
		const start = items[0].start;
		const rect = this.slideEl.getBoundingClientRect();
		const scale = Math.max(this.options.getScale(), 0.05);
		const centerX = rect.left + (start.x + start.w / 2) * scale;
		const centerY = rect.top + (start.y + start.h / 2) * scale;

		this.drag = {
			kind: "rotate",
			pointerId: event.pointerId,
			originX: event.clientX,
			originY: event.clientY,
			items,
			before: this.options.editor.capture([items[0].shape.sourcePart]),
			moved: false,
			frames: new Map(),
			centerX,
			centerY,
			startAngle: Math.atan2(event.clientY - centerY, event.clientX - centerX),
			rotations: new Map(),
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
		switch (drag.kind) {
			case "marquee":
				this.updateMarquee(drag, event);
				return;
			case "rotate":
				this.updateRotate(drag, event);
				return;
			case "move":
				this.updateMove(drag, dxScreen / scale, dyScreen / scale, scale, event);
				return;
			case "resize":
				this.updateResize(drag, dxScreen / scale, dyScreen / scale, scale, event);
				return;
		}
	};

	private onPointerUp = (event: PointerEvent): void => {
		const drag = this.drag;
		if (!drag || event.pointerId !== drag.pointerId) return;
		this.unlisten();
		this.drag = null;
		this.guideLayer?.empty();
		this.marqueeEl = null;

		if (!drag.moved || drag.kind === "marquee") {
			this.syncOverlay();
			return;
		}

		let changed = false;
		for (const item of drag.items) {
			if (!item.shape.source) continue;
			if (drag.kind === "rotate") {
				const rotation = drag.rotations.get(item.shape.id);
				if (rotation === undefined || Math.abs(rotation - item.start.rot) < 0.01) continue;
				writeRotation(item.shape.source, rotation);
				item.shape.frame.rot = rotation;
				changed = true;
				continue;
			}
			const frame = drag.frames.get(item.shape.id);
			if (!frame || sameFrame(frame, item.start)) continue;
			writeFrame(item.shape.source, frame);
			applyToModel(item.shape, frame);
			changed = true;
		}

		// An Alt-drag already recorded the duplicate; the move on top of it is the
		// same gesture, so record it too and let two undos unwind the pair.
		if (changed) {
			const label =
				drag.kind === "rotate"
					? "Rotate"
					: drag.kind === "resize"
						? "Resize"
						: drag.items.length > 1
							? `Move ${drag.items.length} shapes`
							: "Move shape";
			this.options.editor.recordApplied(label, drag.before);
		}
		this.syncOverlay();
	};

	// --------------------------------------------------------------- move

	private updateMove(
		drag: MoveDrag,
		dx: number,
		dy: number,
		scale: number,
		event: PointerEvent,
	): void {
		// Shift constrains to one axis, the way it does in every drawing tool.
		let moveX = event.shiftKey && Math.abs(dx) < Math.abs(dy) ? 0 : dx;
		let moveY = event.shiftKey && Math.abs(dy) <= Math.abs(dx) ? 0 : dy;

		const matched: { xs: number[]; ys: number[] } = { xs: [], ys: [] };
		const snapping = !(event.metaKey || event.ctrlKey) && drag.items.every((i) => !i.start.rot);
		if (snapping) {
			const threshold = SNAP / scale;
			const bounds = unionFrame(drag.items.map((i) => i.start));
			const left = bounds.x + moveX;
			const right = bounds.x + bounds.w + moveX;
			const top = bounds.y + moveY;
			const bottom = bounds.y + bounds.h + moveY;
			const xFit = bestSnap([left, (left + right) / 2, right], drag.guides.xs, threshold);
			if (xFit) {
				moveX += xFit.delta;
				matched.xs.push(xFit.guide);
			}
			const yFit = bestSnap([top, (top + bottom) / 2, bottom], drag.guides.ys, threshold);
			if (yFit) {
				moveY += yFit.delta;
				matched.ys.push(yFit.guide);
			}
		}

		for (const item of drag.items) {
			const frame: Frame = { ...item.start, x: item.start.x + moveX, y: item.start.y + moveY };
			drag.frames.set(item.shape.id, frame);
			applyLive(item.el, item.shape, frame);
		}
		this.drawGuides(matched);
		this.previewOverlay(drag.frames);
	}

	// ------------------------------------------------------------- resize

	private updateResize(
		drag: ResizeDrag,
		dx: number,
		dy: number,
		scale: number,
		event: PointerEvent,
	): void {
		const handle = HANDLES.find((h) => h.id === drag.handle);
		if (!handle) return;
		const single = drag.items.length === 1 ? drag.items[0] : null;

		if (single && single.start.rot) {
			const frame = rotatedResize(single.start, handle, dx, dy, event.shiftKey);
			drag.frames.set(single.shape.id, frame);
			applyLive(single.el, single.shape, frame);
			this.drawGuides({ xs: [], ys: [] });
			this.previewOverlay(drag.frames);
			return;
		}

		const b0 = drag.bounds;
		const movesLeft = handle.fx === 0;
		const movesRight = handle.fx === 1;
		const movesTop = handle.fy === 0;
		const movesBottom = handle.fy === 1;

		let w = b0.w + (movesRight ? dx : movesLeft ? -dx : 0);
		let h = b0.h + (movesBottom ? dy : movesTop ? -dy : 0);
		if (event.shiftKey && movesLeft !== movesRight && movesTop !== movesBottom && b0.h > 0) {
			const ratio = b0.w / b0.h;
			if (Math.abs(w - b0.w) > Math.abs(h - b0.h)) h = w / ratio;
			else w = h * ratio;
		}
		w = Math.max(MIN_SIZE, w);
		h = Math.max(MIN_SIZE, h);

		let x = movesLeft ? b0.x + b0.w - w : b0.x;
		let y = movesTop ? b0.y + b0.h - h : b0.y;

		const matched: { xs: number[]; ys: number[] } = { xs: [], ys: [] };
		if (!(event.metaKey || event.ctrlKey)) {
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
		w = Math.max(MIN_SIZE, w);
		h = Math.max(MIN_SIZE, h);

		// Every shape keeps its place within the box as the box is scaled.
		const sx = b0.w > 0 ? w / b0.w : 1;
		const sy = b0.h > 0 ? h / b0.h : 1;
		for (const item of drag.items) {
			const frame: Frame = {
				...item.start,
				x: x + (item.start.x - b0.x) * sx,
				y: y + (item.start.y - b0.y) * sy,
				w: Math.max(MIN_SIZE, item.start.w * sx),
				h: Math.max(MIN_SIZE, item.start.h * sy),
			};
			drag.frames.set(item.shape.id, frame);
			applyLive(item.el, item.shape, frame);
		}
		this.drawGuides(matched);
		this.previewOverlay(drag.frames);
	}

	// ------------------------------------------------------------- rotate

	private updateRotate(drag: RotateDrag, event: PointerEvent): void {
		const angle = Math.atan2(event.clientY - drag.centerY, event.clientX - drag.centerX);
		const delta = ((angle - drag.startAngle) * 180) / Math.PI;
		for (const item of drag.items) {
			let rotation = item.start.rot + delta;
			if (event.shiftKey) rotation = Math.round(rotation / ROTATE_SNAP) * ROTATE_SNAP;
			rotation = ((rotation % 360) + 360) % 360;
			drag.rotations.set(item.shape.id, rotation);
			item.el.setCssStyles({ transform: `rotate(${rotation}deg)`, transformOrigin: "center center" });
		}
		const box = this.overlay?.querySelector<HTMLElement>(".pptx-selection:not(.is-member)");
		const rotation = drag.rotations.get(drag.items[0].shape.id) ?? 0;
		if (box) box.setCssStyles({ transform: `rotate(${rotation}deg)` });
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
		const extra = this.options.extraGuides?.();
		if (extra) {
			xs.push(...extra.xs);
			ys.push(...extra.ys);
		}
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

	/** Move the outlines with the shapes mid-drag, without a full rebuild. */
	private previewOverlay(frames: Map<string, Frame>): void {
		const overlay = this.overlay;
		if (!overlay) return;
		const items = this.selectedItems();
		const members = Array.from(overlay.querySelectorAll<HTMLElement>(".pptx-selection.is-member"));
		items.forEach((item, i) => {
			const frame = frames.get(item.shape.id);
			const el = members[i];
			if (!frame || !el) return;
			place(el, frame);
		});

		const box = overlay.querySelector<HTMLElement>(".pptx-selection:not(.is-member)");
		if (!box) return;
		const list = items
			.map((i) => frames.get(i.shape.id))
			.filter((f): f is Frame => f !== undefined);
		if (list.length === 0) return;
		const bounds = list.length === 1 ? list[0] : unionFrame(list);
		place(box, bounds);

		const scale = Math.max(this.options.getScale(), 0.05);
		const size = 9 / scale;
		for (const handle of HANDLES) {
			const el = box.querySelector<HTMLElement>(`[data-handle="${handle.id}"]`);
			if (!el) continue;
			el.setCssStyles({
				left: `${bounds.w * handle.fx - size / 2}px`,
				top: `${bounds.h * handle.fy - size / 2}px`,
			});
		}
		const knob = box.querySelector<HTMLElement>("[data-rotate]");
		if (knob) knob.setCssStyles({ left: `${bounds.w / 2 - size / 2}px` });
		const stem = box.querySelector<HTMLElement>(".pptx-rotate-stem");
		if (stem) stem.setCssStyles({ left: `${bounds.w / 2 - 0.75 / scale}px` });
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

		// Tab walks the slide in stacking order, which is the only way to reach a
		// shape sitting exactly behind another one from the keyboard.
		if (event.key === "Tab") {
			const ids = this.shapeIds();
			if (ids.length === 0) return false;
			const current = [...this.options.selection.ids][0];
			const at = current ? ids.indexOf(current) : -1;
			const step = event.shiftKey ? -1 : 1;
			const next = ids[(at + step + ids.length) % ids.length] ?? ids[0];
			this.options.selection.set(this.slideIndex, [next]);
			return true;
		}

		if ((event.key === "Enter" || event.key === "F2") && this.options.selection.size === 1) {
			const id = [...this.options.selection.ids][0];
			if (id) {
				this.options.onEditText?.(id);
				return true;
			}
		}
		if ((event.key === "Delete" || event.key === "Backspace") && ctx) {
			if (this.options.selection.isEmpty) return false;
			deleteSelection(ctx);
			return true;
		}
		// Typing over a selected shape replaces its text, as in PowerPoint. This
		// sits above the modifier bail-out so Shift-typed capitals still count.
		if (
			event.key.length === 1 &&
			!event.metaKey &&
			!event.ctrlKey &&
			!event.altKey &&
			this.options.selection.size === 1
		) {
			const id = [...this.options.selection.ids][0];
			if (id && this.options.onTypeText?.(id, event.key)) return true;
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

		const items = this.dragItems();
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

	private shapeIds(): string[] {
		if (!this.slideEl) return [];
		return Array.from(this.slideEl.querySelectorAll<HTMLElement>("[data-shape-id]"))
			.map((el) => el.dataset.shapeId)
			.filter((id): id is string => Boolean(id));
	}

	/** Select everything on the current slide. */
	selectAll(): void {
		const slideEl = this.slideEl;
		if (!slideEl) return;
		void slideEl;
		this.options.selection.set(this.slideIndex, this.shapeIds());
	}
}

// ------------------------------------------------------------- helpers

function place(el: HTMLElement, frame: Frame): void {
	el.setCssStyles({
		left: `${frame.x}px`,
		top: `${frame.y}px`,
		width: `${frame.w}px`,
		height: `${frame.h}px`,
	});
}

/** Resize a rotated shape along its own axes, holding the anchor corner still. */
function rotatedResize(
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

function applyLive(el: HTMLElement, shape: Shape, frame: Frame): void {
	place(el, frame);
	if (shape.kind === "group") {
		const inner = el.firstElementChild;
		if (inner?.instanceOf(HTMLElement)) {
			const sx = shape.childOffset.w > 0 ? frame.w / shape.childOffset.w : 1;
			const sy = shape.childOffset.h > 0 ? frame.h / shape.childOffset.h : 1;
			inner.setCssStyles({
				transform: `scale(${sx}, ${sy}) translate(${-shape.childOffset.x}px, ${-shape.childOffset.y}px)`,
			});
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

export function unionFrame(frames: Frame[]): Frame {
	const left = Math.min(...frames.map((f) => f.x));
	const top = Math.min(...frames.map((f) => f.y));
	const right = Math.max(...frames.map((f) => f.x + f.w));
	const bottom = Math.max(...frames.map((f) => f.y + f.h));
	return { x: left, y: top, w: right - left, h: bottom - top, rot: 0, flipH: false, flipV: false };
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

function isHtmlElement(value: EventTarget | null): value is HTMLElement {
	return (value as Node | null)?.instanceOf(HTMLElement) === true;
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
