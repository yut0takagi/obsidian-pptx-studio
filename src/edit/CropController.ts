import { setOrderedChild } from "../ooxml/format";
import { relsPathFor } from "../ooxml/rels";
import { A_NS } from "../ooxml/tree";
import type { Crop, Frame, ImageShape, Shape } from "../pptx/types";
import { child } from "../pptx/xml";
import { shapeRegistry } from "../render/renderSlide";
import type { CommandContext } from "./commands";
import { selectedShapes } from "./commands";
import { writeFrame } from "./geometryWrite";

export interface CropControllerOptions {
	getContext: () => CommandContext | null;
	getScale: () => number;
	/** Redraw after entering or leaving crop mode. */
	onChanged: () => void;
}

const HANDLES = [
	{ id: "nw", fx: 0, fy: 0, cursor: "nwse-resize" },
	{ id: "n", fx: 0.5, fy: 0, cursor: "ns-resize" },
	{ id: "ne", fx: 1, fy: 0, cursor: "nesw-resize" },
	{ id: "e", fx: 1, fy: 0.5, cursor: "ew-resize" },
	{ id: "se", fx: 1, fy: 1, cursor: "nwse-resize" },
	{ id: "s", fx: 0.5, fy: 1, cursor: "ns-resize" },
	{ id: "sw", fx: 0, fy: 1, cursor: "nesw-resize" },
	{ id: "w", fx: 0, fy: 0.5, cursor: "ew-resize" },
] as const;

const NO_CROP: Crop = { l: 0, t: 0, r: 0, b: 0 };
/** Never let a crop close down to nothing. */
const MIN_VISIBLE = 0.03;

interface Session {
	shapeId: string;
	/** The whole source image's box on the slide, which crop handles move within. */
	full: { x: number; y: number; w: number; h: number };
	frame: Frame;
	crop: Crop;
}

/**
 * Cropping a picture.
 *
 * A cropped picture is a window onto its source: the shape's frame is the
 * window, and `a:srcRect` says which part of the source it shows. Dragging a
 * crop handle therefore moves both — the frame edge follows the pointer and the
 * inset grows to match, so the part of the image still visible does not shift.
 */
export class CropController {
	private session: Session | null = null;
	private slideEl: HTMLElement | null = null;
	private overlay: HTMLElement | null = null;
	private drag: { pointerId: number; handle: string; start: Session } | null = null;

	constructor(private readonly options: CropControllerOptions) {}

	get active(): boolean {
		return this.session !== null;
	}

	setActive(slideEl: HTMLElement | null): void {
		if (this.slideEl !== slideEl) this.cancel();
		this.slideEl = slideEl;
	}

	/** True when the selection is a single picture, so cropping is offered. */
	canCrop(): boolean {
		return croppableShape(this.options.getContext()) !== null;
	}

	toggle(): void {
		if (this.session) this.finish();
		else this.begin();
	}

	private begin(): void {
		const shape = croppableShape(this.options.getContext());
		if (!shape || !this.slideEl) return;
		const crop = shape.crop ?? NO_CROP;
		this.session = {
			shapeId: shape.id,
			full: fullBox(shape.frame, crop),
			frame: { ...shape.frame },
			crop: { ...crop },
		};
		this.paint();
		this.options.onChanged();
	}

	/** Write the crop into the XML and leave crop mode. */
	finish(): void {
		const session = this.session;
		this.session = null;
		this.clear();
		if (session) {
			const ctx = this.options.getContext();
			const shape = ctx ? findShape(ctx, session.shapeId) : null;
			if (ctx && shape?.source) {
				ctx.editor.transact("Crop picture", [ctx.slide.partPath, relsPathFor(ctx.slide.partPath)], () => {
					writeCrop(shape.source!, session.crop);
					writeFrame(shape.source!, session.frame);
					return true;
				});
			}
		}
		this.options.onChanged();
	}

	cancel(): void {
		if (!this.session) return;
		this.session = null;
		this.clear();
		this.options.onChanged();
	}

	/** Undo the crop entirely, restoring the whole source image. */
	reset(): void {
		const ctx = this.options.getContext();
		const shape = croppableShape(ctx);
		if (!ctx || !shape?.source) return;
		const full = fullBox(shape.frame, shape.crop ?? NO_CROP);
		this.session = null;
		this.clear();
		ctx.editor.transact("Reset crop", [ctx.slide.partPath, relsPathFor(ctx.slide.partPath)], () => {
			writeCrop(shape.source!, NO_CROP);
			writeFrame(shape.source!, { ...shape.frame, ...full });
			return true;
		});
		this.options.onChanged();
	}

	/** Pointer-down while cropping; returns true when a handle was grabbed. */
	tryGrab(event: PointerEvent): boolean {
		if (!this.session) return false;
		const target = event.target;
		const handle =
			target instanceof HTMLElement ? target.closest<HTMLElement>("[data-crop-handle]") : null;
		if (!handle) return false;
		this.drag = {
			pointerId: event.pointerId,
			handle: handle.dataset.cropHandle ?? "",
			start: { ...this.session, crop: { ...this.session.crop }, frame: { ...this.session.frame } },
		};
		window.addEventListener("pointermove", this.onMove);
		window.addEventListener("pointerup", this.onUp);
		event.preventDefault();
		event.stopPropagation();
		return true;
	}

	private onMove = (event: PointerEvent): void => {
		const drag = this.drag;
		const session = this.session;
		if (!drag || !session || event.pointerId !== drag.pointerId || !this.slideEl) return;

		const rect = this.slideEl.getBoundingClientRect();
		const scale = Math.max(this.options.getScale(), 0.05);
		const x = (event.clientX - rect.left) / scale;
		const y = (event.clientY - rect.top) / scale;

		const { full } = drag.start;
		const crop = { ...drag.start.crop };
		const handle = HANDLES.find((h) => h.id === drag.handle);
		if (!handle) return;

		if (handle.fx === 0) crop.l = clamp((x - full.x) / full.w, 0, 1 - crop.r - MIN_VISIBLE);
		if (handle.fx === 1) crop.r = clamp((full.x + full.w - x) / full.w, 0, 1 - crop.l - MIN_VISIBLE);
		if (handle.fy === 0) crop.t = clamp((y - full.y) / full.h, 0, 1 - crop.b - MIN_VISIBLE);
		if (handle.fy === 1) crop.b = clamp((full.y + full.h - y) / full.h, 0, 1 - crop.t - MIN_VISIBLE);

		session.crop = crop;
		session.frame = {
			...drag.start.frame,
			x: full.x + crop.l * full.w,
			y: full.y + crop.t * full.h,
			w: full.w * (1 - crop.l - crop.r),
			h: full.h * (1 - crop.t - crop.b),
		};
		this.paint();
	};

	private onUp = (event: PointerEvent): void => {
		if (!this.drag || event.pointerId !== this.drag.pointerId) return;
		window.removeEventListener("pointermove", this.onMove);
		window.removeEventListener("pointerup", this.onUp);
		this.drag = null;
	};

	private clear(): void {
		this.overlay?.remove();
		this.overlay = null;
	}

	/** Draw the source image dimmed outside the crop window, plus the handles. */
	paint(): void {
		const session = this.session;
		const slideEl = this.slideEl;
		if (!session || !slideEl) return;

		const ctx = this.options.getContext();
		const shape = ctx ? findShape(ctx, session.shapeId) : null;
		if (!shape) return;

		this.clear();
		const overlay = slideEl.createDiv({ cls: "pptx-crop" });
		Object.assign(overlay.style, {
			position: "absolute",
			inset: "0",
			zIndex: "60",
			pointerEvents: "none",
		});
		this.overlay = overlay;

		const { full, frame, crop } = session;
		const ghost = overlay.createDiv({ cls: "pptx-crop-ghost" });
		Object.assign(ghost.style, {
			position: "absolute",
			left: `${full.x}px`,
			top: `${full.y}px`,
			width: `${full.w}px`,
			height: `${full.h}px`,
			opacity: "0.35",
			pointerEvents: "none",
		});
		if (shape.url) {
			const img = ghost.createEl("img");
			img.src = shape.url;
			Object.assign(img.style, { width: "100%", height: "100%", objectFit: "fill" });
		}

		const window_ = overlay.createDiv({ cls: "pptx-crop-window" });
		const scale = Math.max(this.options.getScale(), 0.05);
		const border = 2 / scale;
		Object.assign(window_.style, {
			position: "absolute",
			left: `${frame.x}px`,
			top: `${frame.y}px`,
			width: `${frame.w}px`,
			height: `${frame.h}px`,
			outline: `${border}px solid #ffffff`,
			boxShadow: `0 0 0 ${border}px rgba(0,0,0,0.6)`,
			overflow: "hidden",
			pointerEvents: "none",
		});
		if (shape.url) {
			const img = window_.createEl("img");
			img.src = shape.url;
			Object.assign(img.style, {
				position: "absolute",
				width: `${100 / Math.max(1 - crop.l - crop.r, 0.001)}%`,
				height: `${100 / Math.max(1 - crop.t - crop.b, 0.001)}%`,
				left: `${(-crop.l / Math.max(1 - crop.l - crop.r, 0.001)) * 100}%`,
				top: `${(-crop.t / Math.max(1 - crop.t - crop.b, 0.001)) * 100}%`,
				objectFit: "fill",
			});
		}

		const size = 12 / scale;
		for (const handle of HANDLES) {
			const el = overlay.createDiv({ cls: "pptx-crop-handle" });
			el.dataset.cropHandle = handle.id;
			Object.assign(el.style, {
				position: "absolute",
				pointerEvents: "auto",
				cursor: handle.cursor,
				width: `${size}px`,
				height: `${size}px`,
				left: `${frame.x + frame.w * handle.fx - size / 2}px`,
				top: `${frame.y + frame.h * handle.fy - size / 2}px`,
				background: "#ffffff",
				border: `${1 / scale}px solid rgba(0,0,0,0.6)`,
				borderRadius: `${size / 6}px`,
			});
		}
	}
}

// -------------------------------------------------------------- helpers

/** The selection, when it is exactly one croppable picture. */
function croppableShape(ctx: CommandContext | null): ImageShape | null {
	if (!ctx) return null;
	const shapes = selectedShapes(ctx);
	if (shapes.length !== 1) return null;
	const shape = shapes[0];
	return shape.kind === "image" && shape.source ? shape : null;
}

function findShape(ctx: CommandContext, id: string): ImageShape | null {
	const own = ctx.slide.shapes.slice(ctx.slide.templateShapes);
	const shape: Shape | undefined = own.find((s) => s.id === id);
	return shape?.kind === "image" ? shape : null;
}

/** Where the whole source image would sit, given the visible frame and its crop. */
function fullBox(frame: Frame, crop: Crop): { x: number; y: number; w: number; h: number } {
	const w = frame.w / Math.max(1 - crop.l - crop.r, 0.001);
	const h = frame.h / Math.max(1 - crop.t - crop.b, 0.001);
	return { x: frame.x - crop.l * w, y: frame.y - crop.t * h, w, h };
}

/** Write a:srcRect, removing it when the crop is empty. */
export function writeCrop(source: Element, crop: Crop): boolean {
	const blipFill = child(source, "blipFill");
	if (!blipFill) return false;
	const doc = source.ownerDocument;
	const order = ["blip", "srcRect", "tile", "stretch"];

	const empty = crop.l === 0 && crop.t === 0 && crop.r === 0 && crop.b === 0;
	if (empty) {
		setOrderedChild(blipFill, "srcRect", order, null);
		return true;
	}
	const rect = doc.createElementNS(A_NS, "a:srcRect");
	const pct = (value: number) => String(Math.round(clamp(value, 0, 1) * 100000));
	if (crop.l) rect.setAttribute("l", pct(crop.l));
	if (crop.t) rect.setAttribute("t", pct(crop.t));
	if (crop.r) rect.setAttribute("r", pct(crop.r));
	if (crop.b) rect.setAttribute("b", pct(crop.b));
	setOrderedChild(blipFill, "srcRect", order, rect);
	return true;
}

function clamp(value: number, min: number, max: number): number {
	return Math.max(min, Math.min(max, value));
}

/** Exposed so the shape editor can leave cropping alone while it is on. */
export function isCropTarget(el: HTMLElement): boolean {
	return el.closest("[data-crop-handle]") !== null;
}

export { shapeRegistry };
