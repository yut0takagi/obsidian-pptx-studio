import { setIcon } from "obsidian";
import { t } from "../i18n";
import type { EditController } from "../edit/EditController";
import type { ShapeEditor } from "../edit/ShapeEditor";
import type { PptxPackage } from "../pptx/package";
import type { Deck } from "../pptx/types";
import { renderSlide } from "./renderSlide";

export interface DeckViewerOptions {
	deck: Deck;
	pkg: PptxPackage;
	/** Thumbnail rail and speaker notes are hidden in compact (embed) mode. */
	compact: boolean;
	showThumbnails: boolean;
	showNotes: boolean;
	fitMode: "page" | "width";
	/** Fixed slide to show, disabling navigation. */
	pinnedSlide?: number;
	/**
	 * "toolbar" puts navigation and zoom in a bar above the slide, which is what
	 * embeds want. "status" puts them in a slim bar underneath, leaving the top of
	 * the pane to the ribbon.
	 */
	chrome: "toolbar" | "status" | "none";
	/** Called when the visible slide changes, before the new one is shown. */
	onSlideChange?: (index: number) => void;
	/** Called after a slide is rendered and mounted. */
	onRendered?: (slideEl: HTMLElement | null) => void;
	/** Dragging a thumbnail onto another position. */
	onReorder?: (from: number, to: number) => void;
	/** Right-clicking a thumbnail. */
	onThumbnailMenu?: (index: number, event: MouseEvent) => void;
	/** Zoom, scroll or resize: anything that moves the slide on screen. */
	onViewportChanged?: () => void;
	/** Show rulers around the slide, which is also how guides are created. */
	showRulers?: boolean;
	/** Reserve a column on the right for the selection pane. */
	sidePane?: boolean;
	/** Its width in pixels, and where a drag on its edge reports the new one. */
	sidePaneWidth?: number;
	onSidePaneWidth?: (width: number) => void;
	/** Dragging out of a ruler, with the slide coordinate the drag started at. */
	onRulerDrag?: (
		orientation: "horz" | "vert",
		position: number,
		event: PointerEvent,
	) => void;
	onExportPng?: (slideIndex: number) => void;
	onExtractMarkdown?: () => void;
	onOpenExternal?: () => void;
	onSave?: () => void;
	/** Supplied only where editing is allowed, i.e. the full file view. */
	editor?: EditController;
	shapeEditor?: ShapeEditor;
}

const ZOOM_STEPS = [0.25, 0.35, 0.5, 0.75, 1, 1.25, 1.5, 2, 3];

/** Side pane widths: narrow enough to be a strip, wide enough to read XML in. */
const SIDE_DEFAULT = 260;
const SIDE_MIN = 160;
const SIDE_MAX = 720;

const clampSideWidth = (width: number): number =>
	Math.max(SIDE_MIN, Math.min(SIDE_MAX, Math.round(width)));

/**
 * The deck viewer UI. Slides render once at native size and are scaled with a
 * CSS transform, so zooming and resizing never re-parse or re-layout the deck.
 */
export class DeckViewer {
	private readonly root: HTMLElement;
	private readonly options: DeckViewerOptions;
	private readonly slideCache = new Map<number, HTMLElement>();
	private deck: Deck;

	private stageEl!: HTMLElement;
	private canvasEl!: HTMLElement;
	private railEl: HTMLElement | null = null;
	private notesEl: HTMLElement | null = null;
	private counterEl: HTMLElement | null = null;
	private zoomLabelEl: HTMLElement | null = null;
	private saveButtonEl: HTMLElement | null = null;
	private statusEl: HTMLElement | null = null;
	private sidePaneEl: HTMLElement | null = null;
	private sideResizerEl: HTMLElement | null = null;
	private sideDrag: { pointerId: number; startX: number; startWidth: number } | null = null;
	private rulerH: HTMLElement | null = null;
	private rulerV: HTMLElement | null = null;

	private resizeObserver: ResizeObserver | null = null;
	private thumbObserver: IntersectionObserver | null = null;

	private index = 0;
	/** null means "fit to the pane"; a number is an explicit zoom factor. */
	private zoom: number | null = null;
	private notesVisible: boolean;

	constructor(containerEl: HTMLElement, options: DeckViewerOptions) {
		this.options = options;
		this.deck = options.deck;
		this.notesVisible = options.showNotes && !options.compact;
		this.root = containerEl.createDiv({ cls: "pptx-viewer" });
		if (options.compact) this.root.addClass("is-compact");
		this.index = Math.max(0, Math.min(this.lastIndex, (options.pinnedSlide ?? 1) - 1));
		this.build();
	}

	private get lastIndex(): number {
		return Math.max(0, this.deck.slides.length - 1);
	}

	get currentSlideNumber(): number {
		return this.index + 1;
	}

	private build(): void {
		if (this.options.chrome === "toolbar") this.buildToolbar();

		const body = this.root.createDiv({ cls: "pptx-body" });

		if (this.options.showThumbnails && !this.options.compact && this.deck.slides.length > 1) {
			this.railEl = body.createDiv({ cls: "pptx-rail" });
			this.buildRail();
		}

		const area = body.createDiv({ cls: "pptx-canvas-area" });
		if (this.options.showRulers && !this.options.compact) this.buildRulers(area);
		this.stageEl = area.createDiv({ cls: "pptx-stage" });
		if (this.options.showRulers && !this.options.compact) this.stageEl.addClass("has-rulers");
		this.canvasEl = this.stageEl.createDiv({ cls: "pptx-canvas" });
		this.options.shapeEditor?.attachBackdrop(this.stageEl);

		if (this.options.sidePane && !this.options.compact) {
			this.sideResizerEl = body.createDiv({ cls: "pptx-side-resizer" });
			this.sideResizerEl.addEventListener("pointerdown", this.onSideResizeStart);
			// Double-clicking a divider putting it back where it started is the
			// convention everywhere else that has one.
			this.sideResizerEl.addEventListener("dblclick", () => this.setSideWidth(SIDE_DEFAULT));
			this.sidePaneEl = body.createDiv({ cls: "pptx-side" });
			this.sidePaneEl.style.flexBasis = `${clampSideWidth(this.options.sidePaneWidth ?? SIDE_DEFAULT)}px`;
		}

		if (!this.options.compact) {
			this.notesEl = this.root.createDiv({ cls: "pptx-notes" });
			this.notesEl.toggleClass("is-hidden", !this.notesVisible);
		}
		if (this.options.chrome === "status") this.buildStatusBar();

		this.root.tabIndex = 0;
		this.root.addEventListener("keydown", this.onKeyDown);
		this.stageEl.addEventListener("wheel", this.onWheel, { passive: false });
		this.stageEl.addEventListener("scroll", this.onViewportChanged, { passive: true });

		this.resizeObserver = new ResizeObserver(() => this.applyScale());
		this.resizeObserver.observe(this.stageEl);

		this.showSlide(this.index);
	}

	/** A slim bar under the slide: where it is, how big it is, notes on or off. */
	private buildStatusBar(): void {
		const bar = this.root.createDiv({ cls: "pptx-statusbar" });
		const nav = bar.createDiv({ cls: "pptx-toolbar-group" });
		this.iconButton(nav, "chevron-left", t("nav.previous"), () => this.go(this.index - 1));
		this.counterEl = nav.createSpan({ cls: "pptx-counter" });
		this.iconButton(nav, "chevron-right", t("nav.next"), () => this.go(this.index + 1));

		this.statusEl = bar.createSpan({ cls: "pptx-status" });

		const right = bar.createDiv({ cls: "pptx-toolbar-group pptx-toolbar-right" });
		this.iconButton(right, "sticky-note", t("notes.toggle"), () => this.toggleNotes());
		this.iconButton(right, "zoom-out", t("zoom.out"), () => this.stepZoom(-1));
		this.zoomLabelEl = right.createSpan({ cls: "pptx-zoom-label" });
		this.zoomLabelEl.addEventListener("click", () => this.setZoom(null));
		this.zoomLabelEl.setAttribute("aria-label", t("zoom.fit"));
		this.iconButton(right, "zoom-in", t("zoom.in"), () => this.stepZoom(1));
		this.iconButton(right, "maximize", t("zoom.fit"), () => this.setZoom(null));
	}

	private buildToolbar(): void {
		const bar = this.root.createDiv({ cls: "pptx-toolbar" });
		const nav = bar.createDiv({ cls: "pptx-toolbar-group" });

		if (this.options.pinnedSlide === undefined && this.deck.slides.length > 1) {
			this.iconButton(nav, "chevron-left", t("nav.previous"), () => this.go(this.index - 1));
			this.counterEl = nav.createSpan({ cls: "pptx-counter" });
			this.iconButton(nav, "chevron-right", t("nav.next"), () => this.go(this.index + 1));
		} else {
			this.counterEl = nav.createSpan({ cls: "pptx-counter" });
		}

		const zoomGroup = bar.createDiv({ cls: "pptx-toolbar-group" });
		this.iconButton(zoomGroup, "zoom-out", t("zoom.out"), () => this.stepZoom(-1));
		this.zoomLabelEl = zoomGroup.createSpan({ cls: "pptx-zoom-label" });
		this.zoomLabelEl.addEventListener("click", () => this.setZoom(null));
		this.zoomLabelEl.setAttribute("aria-label", t("zoom.fit"));
		this.iconButton(zoomGroup, "zoom-in", t("zoom.in"), () => this.stepZoom(1));
		this.iconButton(zoomGroup, "maximize", t("zoom.fit"), () => this.setZoom(null));

		const actions = bar.createDiv({ cls: "pptx-toolbar-group pptx-toolbar-right" });
		if (!this.options.compact) {
			this.iconButton(actions, "sticky-note", t("notes.toggle"), () => this.toggleNotes());
		}
		if (this.options.onSave) {
			this.saveButtonEl = this.iconButton(actions, "save", t("cmd.save"), () =>
				this.options.onSave?.(),
			);
			this.saveButtonEl.addClass("pptx-save");
		}
		if (this.options.onExportPng) {
			this.iconButton(actions, "image-down", t("cmd.exportPngTooltip"), () =>
				this.options.onExportPng?.(this.currentSlideNumber),
			);
		}
		if (this.options.onExtractMarkdown) {
			this.iconButton(actions, "file-text", t("cmd.exportMarkdownTooltip"), () =>
				this.options.onExtractMarkdown?.(),
			);
		}
		if (this.options.onOpenExternal) {
			this.iconButton(actions, "external-link", t("view.openExternally"), () =>
				this.options.onOpenExternal?.(),
			);
		}
	}

	private iconButton(
		parent: HTMLElement,
		icon: string,
		label: string,
		onClick: () => void,
	): HTMLElement {
		const btn = parent.createEl("button", { cls: "pptx-btn clickable-icon" });
		setIcon(btn, icon);
		btn.setAttribute("aria-label", label);
		btn.addEventListener("click", (e) => {
			e.preventDefault();
			onClick();
		});
		return btn;
	}

	private buildRail(): void {
		const rail = this.railEl;
		if (!rail) return;
		this.thumbObserver = new IntersectionObserver(
			(entries) => {
				for (const entry of entries) {
					if (!entry.isIntersecting) continue;
					const el = entry.target as HTMLElement;
					this.fillThumbnail(el);
					this.thumbObserver?.unobserve(el);
				}
			},
			{ root: rail, rootMargin: "200px" },
		);

		this.deck.slides.forEach((slide, i) => {
			const item = rail.createDiv({ cls: "pptx-thumb" });
			item.dataset.index = String(i);
			item.setAttribute("aria-label", t("view.slideLabel", { n: slide.index }));
			item.createSpan({ cls: "pptx-thumb-number", text: String(slide.index) });
			item.createDiv({ cls: "pptx-thumb-frame" });
			item.addEventListener("click", () => this.go(i));
			item.addEventListener("contextmenu", (event) => {
				event.preventDefault();
				this.go(i);
				this.options.onThumbnailMenu?.(i, event);
			});
			if (this.options.onReorder) this.makeThumbDraggable(item);
			this.thumbObserver?.observe(item);
		});
	}

	/**
	 * Rulers along the top and left edge.
	 *
	 * They double as the place guides come from: dragging out of a ruler is the
	 * gesture people already know, so no separate "add guide" step is needed.
	 */
	private buildRulers(area: HTMLElement): void {
		this.rulerH = area.createDiv({ cls: "pptx-ruler pptx-ruler-h" });
		this.rulerV = area.createDiv({ cls: "pptx-ruler pptx-ruler-v" });
		for (const [el, orientation] of [
			[this.rulerH, "vert"],
			[this.rulerV, "horz"],
		] as const) {
			el.addEventListener("pointerdown", (event) => {
				const position = this.rulerCoordinate(orientation, event);
				if (position !== null) this.options.onRulerDrag?.(orientation, position, event);
			});
		}
	}

	/** Where in slide coordinates a ruler was pressed. */
	private rulerCoordinate(orientation: "horz" | "vert", event: PointerEvent): number | null {
		const slide = this.canvasEl.firstElementChild as HTMLElement | null;
		if (!slide) return null;
		const rect = slide.getBoundingClientRect();
		const scale = Math.max(this.lastScale, 0.05);
		return orientation === "vert"
			? (event.clientX - rect.left) / scale
			: (event.clientY - rect.top) / scale;
	}

	/** Redraw ruler ticks for the current zoom and slide position. */
	private paintRulers(): void {
		const slide = this.canvasEl.firstElementChild as HTMLElement | null;
		if (!slide || !this.rulerH || !this.rulerV) return;
		const scale = Math.max(this.lastScale, 0.05);
		const slideRect = slide.getBoundingClientRect();

		// Choose a spacing that keeps major ticks at least 60 screen pixels apart.
		const candidates = [10, 20, 25, 50, 100, 200, 250, 500, 1000];
		const step = candidates.find((c) => c * scale >= 60) ?? 1000;

		for (const [el, horizontal, size] of [
			[this.rulerH, true, this.deck.width],
			[this.rulerV, false, this.deck.height],
		] as const) {
			const rect = el.getBoundingClientRect();
			const origin = horizontal ? slideRect.left - rect.left : slideRect.top - rect.top;
			el.empty();
			for (let value = 0; value <= size + 0.5; value += step) {
				const at = origin + value * scale;
				if (at < -20 || at > (horizontal ? rect.width : rect.height) + 20) continue;
				const tick = el.createDiv({ cls: "pptx-ruler-tick" });
				tick.style[horizontal ? "left" : "top"] = `${at}px`;
				tick.createSpan({ cls: "pptx-ruler-label", text: String(value) });
			}
		}
	}

	/**
	 * Drag a thumbnail to reorder the deck. The drop position is the gap the
	 * pointer is nearest, which is what the insertion line drawn during the drag
	 * is showing.
	 */
	private makeThumbDraggable(item: HTMLElement): void {
		item.draggable = true;
		item.addEventListener("dragstart", (event) => {
			event.dataTransfer?.setData("text/plain", item.dataset.index ?? "");
			if (event.dataTransfer) event.dataTransfer.effectAllowed = "move";
			item.addClass("is-dragging");
		});
		item.addEventListener("dragend", () => {
			item.removeClass("is-dragging");
			this.clearDropMarkers();
		});
		item.addEventListener("dragover", (event) => {
			event.preventDefault();
			if (event.dataTransfer) event.dataTransfer.dropEffect = "move";
			const rect = item.getBoundingClientRect();
			const after = event.clientY > rect.top + rect.height / 2;
			this.clearDropMarkers();
			item.addClass(after ? "is-drop-after" : "is-drop-before");
		});
		item.addEventListener("dragleave", () => this.clearDropMarkers());
		item.addEventListener("drop", (event) => {
			event.preventDefault();
			const from = Number(event.dataTransfer?.getData("text/plain"));
			const target = Number(item.dataset.index);
			const rect = item.getBoundingClientRect();
			const after = event.clientY > rect.top + rect.height / 2;
			this.clearDropMarkers();
			if (!Number.isFinite(from) || !Number.isFinite(target)) return;
			// Dropping below a thumbnail means "after it", which is one index later
			// unless the slide is moving down the list and vacates its own slot.
			let to = after ? target + 1 : target;
			if (from < to) to -= 1;
			if (to === from) return;
			this.options.onReorder?.(from, to);
		});
	}

	private clearDropMarkers(): void {
		if (!this.railEl) return;
		for (const el of Array.from(this.railEl.children)) {
			(el as HTMLElement).removeClass("is-drop-before");
			(el as HTMLElement).removeClass("is-drop-after");
		}
	}

	/** Thumbnails render only once they scroll into view; big decks stay responsive. */
	private fillThumbnail(item: HTMLElement): void {
		const index = Number(item.dataset.index);
		const frame = item.querySelector<HTMLElement>(".pptx-thumb-frame");
		if (!frame || frame.hasChildNodes()) return;
		if (!frame.isConnected) return;
		const slide = this.deck.slides[index];
		if (!slide) return;

		const width = frame.clientWidth || 160;
		const scale = width / this.deck.width;
		frame.style.height = `${this.deck.height * scale}px`;

		const el = renderSlide(this.deck, slide);
		el.style.transform = `scale(${scale})`;
		el.style.pointerEvents = "none";
		frame.appendChild(el);
	}

	private slideElement(index: number): HTMLElement | null {
		const slide = this.deck.slides[index];
		if (!slide) return null;
		const cached = this.slideCache.get(index);
		if (cached) return cached;
		const el = renderSlide(this.deck, slide, { border: true });
		this.options.editor?.attach(el);
		this.options.shapeEditor?.attach(el);
		this.slideCache.set(index, el);
		return el;
	}

	private showSlide(index: number): void {
		// Repainting the current slide after an edit is not a slide change.
		// Reporting it as one is what used to drop the selection every time a
		// ribbon command ran — press "centre", lose the shape you centred. The
		// slide's own part path is the identity, so deleting or reordering
		// slides still counts as a change even when the index stays put.
		const identity = this.deck.slides[index]?.partPath ?? null;
		if (identity !== this.notifiedSlide) {
			this.notifiedSlide = identity;
			this.options.onSlideChange?.(index);
		}
		const el = this.slideElement(index);
		this.canvasEl.empty();
		if (el) this.canvasEl.appendChild(el);
		// The slide element's own transform belongs to the zoom, so the entry
		// animation moves the canvas around it rather than overwriting it.
		this.canvasEl.removeClass("is-enter-forward");
		this.canvasEl.removeClass("is-enter-back");
		if (this.navigationDirection && el) {
			const cls = this.navigationDirection === "forward" ? "is-enter-forward" : "is-enter-back";
			// Reading offsetWidth restarts the animation when the class is re-added.
			void this.canvasEl.offsetWidth;
			this.canvasEl.addClass(cls);
		}
		this.options.shapeEditor?.setActive(el);
		this.options.onRendered?.(el);

		if (this.counterEl) {
			this.counterEl.setText(
				this.deck.slides.length ? `${index + 1} / ${this.deck.slides.length}` : "0 / 0",
			);
		}
		if (this.railEl) {
			for (const item of Array.from(this.railEl.children)) {
				item.toggleClass("is-active", Number((item as HTMLElement).dataset.index) === index);
			}
			const active = this.railEl.querySelector<HTMLElement>(".pptx-thumb.is-active");
			active?.scrollIntoView({ block: "nearest", behavior: "smooth" });
		}
		if (this.notesEl) {
			const notes = this.deck.slides[index]?.notes ?? "";
			this.notesEl.empty();
			this.notesEl.createDiv({ cls: "pptx-notes-title", text: t("view.notesTitle") });
			this.notesEl.createDiv({
				cls: "pptx-notes-body",
				text: notes || t("view.noNotes"),
			});
			this.notesEl.toggleClass("is-empty", !notes);
		}
		this.applyScale();
	}

	private applyScale(): void {
		const el = this.canvasEl.firstElementChild as HTMLElement | null;
		if (!el) return;
		const available = this.stageEl.getBoundingClientRect();
		const padding = this.options.compact ? 0 : 24;
		const availableWidth = Math.max(40, available.width - padding);
		const availableHeight = Math.max(40, available.height - padding);

		let scale: number;
		if (this.zoom !== null) {
			scale = this.zoom;
		} else if (this.options.fitMode === "width") {
			scale = availableWidth / this.deck.width;
		} else {
			scale = Math.min(availableWidth / this.deck.width, availableHeight / this.deck.height);
		}
		scale = Math.max(0.05, scale);

		el.style.transform = `scale(${scale})`;
		this.canvasEl.style.width = `${this.deck.width * scale}px`;
		this.canvasEl.style.height = `${this.deck.height * scale}px`;
		this.zoomLabelEl?.setText(`${Math.round(scale * 100)}%`);
		this.lastScale = scale;
		this.options.shapeEditor?.refresh();
		this.paintRulers();
		this.options.onViewportChanged?.();
	}

	/** The factor the slide element is currently scaled by. */
	currentScale(): number {
		return this.lastScale;
	}

	private lastScale = 1;

	private onViewportChanged = (): void => {
		this.options.onViewportChanged?.();
	};

	/** The scrolling box the slide sits in, for positioning floating UI. */
	get stageElement(): HTMLElement | null {
		return this.stageEl ?? null;
	}

	/** A short line about the selection, shown next to the slide counter. */
	setStatus(text: string): void {
		this.statusEl?.setText(text);
	}

	/** Where the selection pane mounts, when one was asked for. */
	get sidePane(): HTMLElement | null {
		return this.sidePaneEl;
	}

	/** Set the side pane's width and re-fit the slide into what is left. */
	private setSideWidth(width: number): void {
		if (!this.sidePaneEl) return;
		const clamped = clampSideWidth(width);
		this.sidePaneEl.style.flexBasis = `${clamped}px`;
		this.applyScale();
		this.options.onViewportChanged?.();
	}

	private onSideResizeStart = (event: PointerEvent): void => {
		if (event.button !== 0 || !this.sidePaneEl) return;
		event.preventDefault();
		this.sideDrag = {
			pointerId: event.pointerId,
			startX: event.clientX,
			startWidth: this.sidePaneEl.getBoundingClientRect().width,
		};
		this.sideResizerEl?.addClass("is-dragging");
		window.addEventListener("pointermove", this.onSideResizeMove);
		window.addEventListener("pointerup", this.onSideResizeEnd);
		window.addEventListener("pointercancel", this.onSideResizeEnd);
	};

	private onSideResizeMove = (event: PointerEvent): void => {
		const drag = this.sideDrag;
		if (!drag || event.pointerId !== drag.pointerId) return;
		// The pane is on the right, so pulling the divider left widens it.
		this.setSideWidth(drag.startWidth - (event.clientX - drag.startX));
	};

	private onSideResizeEnd = (event: PointerEvent): void => {
		const drag = this.sideDrag;
		if (!drag || event.pointerId !== drag.pointerId) return;
		this.sideDrag = null;
		this.sideResizerEl?.removeClass("is-dragging");
		window.removeEventListener("pointermove", this.onSideResizeMove);
		window.removeEventListener("pointerup", this.onSideResizeEnd);
		window.removeEventListener("pointercancel", this.onSideResizeEnd);
		const width = this.sidePaneEl?.getBoundingClientRect().width;
		if (width) this.options.onSidePaneWidth?.(Math.round(width));
	};

	/** Slide count, for callers driving navigation from outside. */
	get slideCount(): number {
		return this.deck.slides.length;
	}

	next(): void {
		this.go(this.index + 1);
	}

	previous(): void {
		this.go(this.index - 1);
	}

	zoomIn(): void {
		this.stepZoom(1);
	}

	zoomOut(): void {
		this.stepZoom(-1);
	}

	zoomToFit(): void {
		this.setZoom(null);
	}

	showNotes(visible: boolean): void {
		if (this.notesVisible === visible) return;
		this.toggleNotes();
	}

	get notesShown(): boolean {
		return this.notesVisible;
	}

	private stepZoom(direction: number): void {
		const current = this.currentScale();
		const next =
			direction > 0
				? (ZOOM_STEPS.find((z) => z > current + 0.001) ?? ZOOM_STEPS[ZOOM_STEPS.length - 1])
				: ([...ZOOM_STEPS].reverse().find((z) => z < current - 0.001) ?? ZOOM_STEPS[0]);
		this.setZoom(next);
	}

	private setZoom(value: number | null): void {
		this.zoom = value;
		this.applyScale();
	}

	private toggleNotes(): void {
		this.notesVisible = !this.notesVisible;
		this.notesEl?.toggleClass("is-hidden", !this.notesVisible);
		this.applyScale();
	}

	go(index: number): void {
		if (this.options.pinnedSlide !== undefined) return;
		const clamped = Math.max(0, Math.min(this.lastIndex, index));
		if (clamped === this.index && this.canvasEl.hasChildNodes()) return;
		// Only navigation animates. An edit re-rendering the same slide must not
		// replay the transition, or typing would strobe.
		this.navigationDirection = clamped > this.index ? "forward" : "back";
		this.index = clamped;
		this.showSlide(clamped);
		this.navigationDirection = null;
	}

	private navigationDirection: "forward" | "back" | null = null;
	/** The slide `onSlideChange` last reported, by part path. */
	private notifiedSlide: string | null = null;

	private onKeyDown = (event: KeyboardEvent): void => {
		// A selected shape claims the arrow keys for nudging before paging sees them.
		if (this.options.shapeEditor?.handleKey(event)) {
			event.preventDefault();
			event.stopPropagation();
			return;
		}
		switch (event.key) {
			// Slides are stacked vertically in the rail, so they are paged the way
			// they are stacked: down is the next one.
			case "ArrowDown":
			case "PageDown":
			case "j":
				this.go(this.index + 1);
				break;
			case "ArrowUp":
			case "PageUp":
			case "k":
				this.go(this.index - 1);
				break;
			case "Home":
				this.go(0);
				break;
			case "End":
				this.go(this.lastIndex);
				break;
			case "+":
			case "=":
				this.stepZoom(1);
				break;
			case "-":
				this.stepZoom(-1);
				break;
			case "0":
				this.setZoom(null);
				break;
			case "n":
				if (!this.options.compact) this.toggleNotes();
				break;
			default:
				return;
		}
		event.preventDefault();
		event.stopPropagation();
	};

	/** Ctrl/Cmd + wheel zooms; a plain wheel keeps the pane's normal scrolling. */
	private onWheel = (event: WheelEvent): void => {
		if (!event.ctrlKey && !event.metaKey) return;
		event.preventDefault();
		const current = this.currentScale();
		this.setZoom(Math.max(0.05, Math.min(6, current * (event.deltaY < 0 ? 1.1 : 1 / 1.1))));
	};

	focus(): void {
		this.root.focus();
	}

	/**
	 * Swap in a rebuilt deck after an edit and redraw. Cached slide elements are
	 * dropped because they hold registry entries pointing at the previous model.
	 */
	setDeck(deck: Deck): void {
		const grew = deck.slides.length > this.deck.slides.length;
		this.deck = deck;
		this.rebuildRail();
		// A slide that just appeared is worth pointing at; the rest of the rail
		// redrawing is not, so only the new one animates.
		if (grew) this.markNewThumbnail(this.index);
		this.index = Math.max(0, Math.min(this.lastIndex, this.index));
		this.invalidate();
	}

	/**
	 * Refresh one slide after an edit that only touched it.
	 *
	 * The alternative — dropping every cached slide and re-rendering — is what
	 * made a keystroke feel like a page load on a long deck.
	 */
	refreshSlide(index: number): void {
		const cached = this.slideCache.get(index);
		if (cached) {
			this.options.editor?.detach(cached);
			this.options.shapeEditor?.detach(cached);
			this.slideCache.delete(index);
		}
		this.refreshThumbnail(index);
		if (index === this.index) this.showSlide(index);
	}

	private markNewThumbnail(index: number): void {
		const item = this.railEl?.children[index] as HTMLElement | undefined;
		if (!item) return;
		item.addClass("is-new");
		window.setTimeout(() => item.removeClass("is-new"), 600);
	}

	/** Re-render one thumbnail, leaving the rest of the rail alone. */
	private refreshThumbnail(index: number): void {
		const item = this.railEl?.children[index] as HTMLElement | undefined;
		const frame = item?.querySelector<HTMLElement>(".pptx-thumb-frame");
		if (!frame) return;
		frame.empty();
		this.fillThumbnail(item as HTMLElement);
	}

	/** Rebuild the whole rail, for when slides were added, removed or reordered. */
	private rebuildRail(): void {
		if (!this.railEl) return;
		this.thumbObserver?.disconnect();
		this.thumbObserver = null;
		this.railEl.empty();
		this.buildRail();
	}

	/** Re-render from the current model, discarding cached slide elements. */
	invalidate(): void {
		for (const el of this.slideCache.values()) {
			this.options.editor?.detach(el);
			this.options.shapeEditor?.detach(el);
		}
		this.slideCache.clear();
		this.showSlide(this.index);
	}

	setDirty(dirty: boolean): void {
		this.saveButtonEl?.toggleClass("is-dirty", dirty);
	}

	destroy(): void {
		for (const el of this.slideCache.values()) {
			this.options.editor?.detach(el);
			this.options.shapeEditor?.detach(el);
		}
		this.resizeObserver?.disconnect();
		this.thumbObserver?.disconnect();
		this.root.removeEventListener("keydown", this.onKeyDown);
		window.removeEventListener("pointermove", this.onSideResizeMove);
		window.removeEventListener("pointerup", this.onSideResizeEnd);
		window.removeEventListener("pointercancel", this.onSideResizeEnd);
		if (this.stageEl) this.options.shapeEditor?.detachBackdrop(this.stageEl);
		this.stageEl?.removeEventListener("wheel", this.onWheel);
		this.stageEl?.removeEventListener("scroll", this.onViewportChanged);
		this.slideCache.clear();
		this.root.detach();
	}
}
