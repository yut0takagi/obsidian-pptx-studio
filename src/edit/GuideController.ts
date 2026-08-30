import type { Guide } from "../ooxml/guides";
import { guideParts, readGuides, writeGuides } from "../ooxml/guides";
import type { PptxPackage } from "../pptx/package";
import type { DeckEditor } from "./DeckEditor";

export interface GuideControllerOptions {
	pkg: PptxPackage;
	editor: DeckEditor;
	getScale: () => number;
	/** Slide size in pixels, so a guide dragged off the edge is deleted. */
	getSlideSize: () => { width: number; height: number };
	isEnabled: () => boolean;
}

const HIT_TOLERANCE = 4;

/**
 * Drawing guides: drag one out of a ruler, drag it to move, drag it off the
 * slide to remove it.
 *
 * The guides are read from and written back to the deck, so they survive a save
 * and appear in PowerPoint too. They are not part of any slide's content, which
 * is why they live outside the shape editor entirely.
 */
export class GuideController {
	private guides: Guide[] = [];
	private slideEl: HTMLElement | null = null;
	private layer: HTMLElement | null = null;
	private drag: {
		pointerId: number;
		index: number;
		orientation: Guide["orientation"];
		/** Guides created by dragging out of a ruler start life uncommitted. */
		created: boolean;
	} | null = null;

	constructor(private readonly options: GuideControllerOptions) {
		this.guides = readGuides(options.pkg);
	}

	/** Re-read after an undo or a reload replaced the parts. */
	reload(): void {
		this.guides = readGuides(this.options.pkg);
		this.paint();
	}

	get all(): readonly Guide[] {
		return this.guides;
	}

	/** Guide positions in the form the shape editor snaps against. */
	snapTargets(): { xs: number[]; ys: number[] } {
		return {
			xs: this.guides.filter((g) => g.orientation === "vert").map((g) => g.position),
			ys: this.guides.filter((g) => g.orientation === "horz").map((g) => g.position),
		};
	}

	setActive(slideEl: HTMLElement | null): void {
		this.layer = null;
		this.slideEl = slideEl;
		this.paint();
	}

	/** Start dragging a new guide out of a ruler. */
	beginFromRuler(event: PointerEvent, orientation: Guide["orientation"], position: number): void {
		if (!this.options.isEnabled()) return;
		this.guides.push({ orientation, position });
		this.drag = {
			pointerId: event.pointerId,
			index: this.guides.length - 1,
			orientation,
			created: true,
		};
		this.paint();
		this.listen();
		event.preventDefault();
	}

	/** Pointer-down inside the slide; returns true when it grabbed a guide. */
	tryGrab(event: PointerEvent): boolean {
		if (!this.options.isEnabled() || !this.slideEl) return false;
		const point = this.toSlide(event);
		if (!point) return false;
		const tolerance = HIT_TOLERANCE / Math.max(this.options.getScale(), 0.05);

		const index = this.guides.findIndex((guide) =>
			guide.orientation === "vert"
				? Math.abs(guide.position - point.x) <= tolerance
				: Math.abs(guide.position - point.y) <= tolerance,
		);
		if (index === -1) return false;

		this.drag = {
			pointerId: event.pointerId,
			index,
			orientation: this.guides[index].orientation,
			created: false,
		};
		this.listen();
		event.preventDefault();
		event.stopPropagation();
		return true;
	}

	private listen(): void {
		window.addEventListener("pointermove", this.onMove);
		window.addEventListener("pointerup", this.onUp);
		window.addEventListener("pointercancel", this.onUp);
	}

	private onMove = (event: PointerEvent): void => {
		const drag = this.drag;
		if (!drag || event.pointerId !== drag.pointerId) return;
		const point = this.toSlide(event);
		if (!point) return;
		const guide = this.guides[drag.index];
		if (!guide) return;
		guide.position = drag.orientation === "vert" ? point.x : point.y;
		this.paint();
	};

	private onUp = (event: PointerEvent): void => {
		const drag = this.drag;
		if (!drag || event.pointerId !== drag.pointerId) return;
		window.removeEventListener("pointermove", this.onMove);
		window.removeEventListener("pointerup", this.onUp);
		window.removeEventListener("pointercancel", this.onUp);
		this.drag = null;

		// Dropped outside the slide means "throw it away", the way it does in
		// every other tool with rulers.
		const guide = this.guides[drag.index];
		const { width, height } = this.options.getSlideSize();
		const limit = guide?.orientation === "vert" ? width : height;
		if (!guide || guide.position < 0 || guide.position > limit) {
			this.guides.splice(drag.index, 1);
		} else {
			guide.position = Math.round(guide.position);
		}

		this.commit();
		this.paint();
	};

	private commit(): void {
		const snapshot = [...this.guides.map((g) => ({ ...g }))];
		this.options.editor.transact(
			"Guides",
			guideParts(),
			() => writeGuides(this.options.pkg, snapshot),
			// Guides are not slide content, so nothing about the model changed.
			{ rebuild: false },
		);
	}

	private toSlide(event: PointerEvent): { x: number; y: number } | null {
		const el = this.slideEl;
		if (!el) return null;
		const rect = el.getBoundingClientRect();
		const scale = Math.max(this.options.getScale(), 0.05);
		return { x: (event.clientX - rect.left) / scale, y: (event.clientY - rect.top) / scale };
	}

	/** Draw the guides into the active slide element. */
	paint(): void {
		const slideEl = this.slideEl;
		if (!slideEl) return;
		if (!this.layer?.isConnected) {
			this.layer = slideEl.createDiv({ cls: "pptx-user-guides" });
			Object.assign(this.layer.style, {
				position: "absolute",
				inset: "0",
				pointerEvents: "none",
				zIndex: "30",
			});
		}
		const layer = this.layer;
		layer.empty();

		const scale = Math.max(this.options.getScale(), 0.05);
		const thickness = 1 / scale;
		for (const guide of this.guides) {
			const line = layer.createDiv({ cls: "pptx-user-guide" });
			Object.assign(line.style, {
				position: "absolute",
				background: "rgba(220, 60, 190, 0.75)",
			});
			if (guide.orientation === "vert") {
				Object.assign(line.style, {
					left: `${guide.position - thickness / 2}px`,
					top: "0",
					width: `${thickness}px`,
					height: "100%",
				});
			} else {
				Object.assign(line.style, {
					top: `${guide.position - thickness / 2}px`,
					left: "0",
					height: `${thickness}px`,
					width: "100%",
				});
			}
		}
	}

	/** Add a guide down the middle of the slide, for the menu command. */
	addCentre(orientation: Guide["orientation"]): void {
		const { width, height } = this.options.getSlideSize();
		this.guides.push({
			orientation,
			position: Math.round((orientation === "vert" ? width : height) / 2),
		});
		this.commit();
		this.paint();
	}

	clearAll(): void {
		if (this.guides.length === 0) return;
		this.guides = [];
		this.commit();
		this.paint();
	}
}
