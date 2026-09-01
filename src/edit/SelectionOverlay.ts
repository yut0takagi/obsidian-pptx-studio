/**
 * Everything drawn *over* the slide: the selection outlines and their handles,
 * the snapping guides, and the marquee.
 *
 * None of it belongs to the deck, so none of it is ever written back. It lives
 * in two layers of its own above the rendered shapes, and it is sized in slide
 * coordinates divided by the view's zoom — a handle has to stay the same size
 * on screen whatever the slide is scaled to, or it becomes unclickable at a
 * fit-to-window zoom.
 */
import type { Frame, Shape } from "../pptx/types";
import { HANDLES, unionFrame } from "./dragMath";
import { place } from "./liveFrame";

const ACCENT = "var(--interactive-accent, #2f6fed)";

export interface OverlayItem {
	shape: Shape;
	el: HTMLElement;
}

export class SelectionOverlay {
	private slideEl: HTMLElement | null = null;
	private overlay: HTMLElement | null = null;
	private guideLayer: HTMLElement | null = null;
	private marqueeEl: HTMLElement | null = null;

	constructor(private readonly getScale: () => number) {}

	/** Which rendered slide the layers belong over. */
	setSlide(slideEl: HTMLElement | null): void {
		this.clear();
		this.slideEl = slideEl;
	}

	get slide(): HTMLElement | null {
		return this.slideEl;
	}

	clear(): void {
		this.overlay?.remove();
		this.guideLayer?.remove();
		this.overlay = null;
		this.guideLayer = null;
		this.marqueeEl = null;
	}

	/** Drop the guides and the marquee, keeping the selection drawn. */
	endGesture(): void {
		this.guideLayer?.empty();
		this.marqueeEl = null;
	}

	private ensureLayers(): void {
		const slideEl = this.slideEl;
		if (!slideEl) return;
		if (this.guideLayer?.isConnected && this.overlay?.isConnected) return;
		this.clear();

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

	/** Slide scale, floored so a division by it can never blow up. */
	private get scale(): number {
		return Math.max(this.getScale(), 0.05);
	}

	sync(items: OverlayItem[], editing: boolean): void {
		if (!this.slideEl) return;
		this.ensureLayers();
		const overlay = this.overlay;
		if (!overlay) return;
		overlay.empty();
		if (items.length === 0) return;

		const scale = this.scale;
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
					outline: `${border}px dashed ${ACCENT}`,
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
			outline: `${border}px ${editing ? "dashed" : "solid"} ${ACCENT}`,
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
				border: `${border}px solid ${ACCENT}`,
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
				background: ACCENT,
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
				border: `${border}px solid ${ACCENT}`,
				background: "var(--background-primary, #fff)",
				borderRadius: "50%",
			});
		}
	}

	/** Move the outlines with the shapes mid-drag, without a full rebuild. */
	preview(items: OverlayItem[], frames: Map<string, Frame>): void {
		const overlay = this.overlay;
		if (!overlay) return;
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

		const scale = this.scale;
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

	/** Spin the selection box with a rotate drag, before anything is written. */
	previewRotation(degrees: number): void {
		const box = this.overlay?.querySelector<HTMLElement>(".pptx-selection:not(.is-member)");
		if (box) box.setCssStyles({ transform: `rotate(${degrees}deg)` });
	}

	/** Draw the guides a drag has snapped to, replacing whatever was there. */
	showGuides(matched: { xs: number[]; ys: number[] }): void {
		const layer = this.guideLayer;
		if (!layer) return;
		for (const el of Array.from(layer.querySelectorAll(".pptx-guide"))) el.remove();
		const thickness = 1 / this.scale;
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

	showMarquee(box: { x: number; y: number; w: number; h: number }): void {
		this.ensureLayers();
		if (!this.marqueeEl && this.guideLayer) {
			this.marqueeEl = this.guideLayer.createDiv({ cls: "pptx-marquee" });
			Object.assign(this.marqueeEl.style, {
				position: "absolute",
				border: `${1 / this.scale}px solid ${ACCENT}`,
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
	}
}
