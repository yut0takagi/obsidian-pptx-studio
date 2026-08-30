import { setIcon } from "obsidian";
import type { EditController } from "../edit/EditController";
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
	/** Show the toolbar. */
	controls: boolean;
	onExportPng?: (slideIndex: number) => void;
	onExtractMarkdown?: () => void;
	onOpenExternal?: () => void;
	onSave?: () => void;
	/** Supplied only where editing is allowed, i.e. the full file view. */
	editor?: EditController;
}

const ZOOM_STEPS = [0.25, 0.35, 0.5, 0.75, 1, 1.25, 1.5, 2, 3];

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
		if (this.options.controls) this.buildToolbar();

		const body = this.root.createDiv({ cls: "pptx-body" });

		if (this.options.showThumbnails && !this.options.compact && this.deck.slides.length > 1) {
			this.railEl = body.createDiv({ cls: "pptx-rail" });
			this.buildRail();
		}

		this.stageEl = body.createDiv({ cls: "pptx-stage" });
		this.canvasEl = this.stageEl.createDiv({ cls: "pptx-canvas" });

		if (!this.options.compact) {
			this.notesEl = this.root.createDiv({ cls: "pptx-notes" });
			this.notesEl.toggleClass("is-hidden", !this.notesVisible);
		}

		this.root.tabIndex = 0;
		this.root.addEventListener("keydown", this.onKeyDown);
		this.stageEl.addEventListener("wheel", this.onWheel, { passive: false });

		this.resizeObserver = new ResizeObserver(() => this.applyScale());
		this.resizeObserver.observe(this.stageEl);

		this.showSlide(this.index);
	}

	private buildToolbar(): void {
		const bar = this.root.createDiv({ cls: "pptx-toolbar" });
		const nav = bar.createDiv({ cls: "pptx-toolbar-group" });

		if (this.options.pinnedSlide === undefined && this.deck.slides.length > 1) {
			this.iconButton(nav, "chevron-left", "Previous slide (←)", () => this.go(this.index - 1));
			this.counterEl = nav.createSpan({ cls: "pptx-counter" });
			this.iconButton(nav, "chevron-right", "Next slide (→)", () => this.go(this.index + 1));
		} else {
			this.counterEl = nav.createSpan({ cls: "pptx-counter" });
		}

		const zoomGroup = bar.createDiv({ cls: "pptx-toolbar-group" });
		this.iconButton(zoomGroup, "zoom-out", "Zoom out (−)", () => this.stepZoom(-1));
		this.zoomLabelEl = zoomGroup.createSpan({ cls: "pptx-zoom-label" });
		this.zoomLabelEl.addEventListener("click", () => this.setZoom(null));
		this.zoomLabelEl.setAttribute("aria-label", "Fit to pane (0)");
		this.iconButton(zoomGroup, "zoom-in", "Zoom in (+)", () => this.stepZoom(1));
		this.iconButton(zoomGroup, "maximize", "Fit to pane (0)", () => this.setZoom(null));

		const actions = bar.createDiv({ cls: "pptx-toolbar-group pptx-toolbar-right" });
		if (!this.options.compact) {
			this.iconButton(actions, "sticky-note", "Toggle speaker notes (N)", () => this.toggleNotes());
		}
		if (this.options.onSave) {
			this.saveButtonEl = this.iconButton(actions, "save", "Save the deck (Cmd/Ctrl+S)", () =>
				this.options.onSave?.(),
			);
			this.saveButtonEl.addClass("pptx-save");
		}
		if (this.options.onExportPng) {
			this.iconButton(actions, "image-down", "Export this slide as PNG", () =>
				this.options.onExportPng?.(this.currentSlideNumber),
			);
		}
		if (this.options.onExtractMarkdown) {
			this.iconButton(actions, "file-text", "Extract deck text to Markdown", () =>
				this.options.onExtractMarkdown?.(),
			);
		}
		if (this.options.onOpenExternal) {
			this.iconButton(actions, "external-link", "Open in the default app", () =>
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
			item.setAttribute("aria-label", `Slide ${slide.index}`);
			item.createSpan({ cls: "pptx-thumb-number", text: String(slide.index) });
			item.createDiv({ cls: "pptx-thumb-frame" });
			item.addEventListener("click", () => this.go(i));
			this.thumbObserver?.observe(item);
		});
	}

	/** Thumbnails render only once they scroll into view; big decks stay responsive. */
	private fillThumbnail(item: HTMLElement): void {
		const index = Number(item.dataset.index);
		const frame = item.querySelector<HTMLElement>(".pptx-thumb-frame");
		if (!frame || frame.hasChildNodes()) return;
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
		this.slideCache.set(index, el);
		return el;
	}

	private showSlide(index: number): void {
		const el = this.slideElement(index);
		this.canvasEl.empty();
		if (el) this.canvasEl.appendChild(el);

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
			active?.scrollIntoView({ block: "nearest" });
		}
		if (this.notesEl) {
			const notes = this.deck.slides[index]?.notes ?? "";
			this.notesEl.empty();
			this.notesEl.createDiv({ cls: "pptx-notes-title", text: "Speaker notes" });
			this.notesEl.createDiv({
				cls: "pptx-notes-body",
				text: notes || "This slide has no speaker notes.",
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
	}

	private stepZoom(direction: number): void {
		const current = this.currentScale();
		const next =
			direction > 0
				? (ZOOM_STEPS.find((z) => z > current + 0.001) ?? ZOOM_STEPS[ZOOM_STEPS.length - 1])
				: ([...ZOOM_STEPS].reverse().find((z) => z < current - 0.001) ?? ZOOM_STEPS[0]);
		this.setZoom(next);
	}

	private currentScale(): number {
		if (this.zoom !== null) return this.zoom;
		const el = this.canvasEl.firstElementChild as HTMLElement | null;
		if (!el) return 1;
		const match = /scale\(([\d.]+)\)/.exec(el.style.transform);
		return match ? Number(match[1]) : 1;
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
		this.index = clamped;
		this.showSlide(clamped);
	}

	private onKeyDown = (event: KeyboardEvent): void => {
		switch (event.key) {
			case "ArrowRight":
			case "PageDown":
			case "j":
				this.go(this.index + 1);
				break;
			case "ArrowLeft":
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
		this.deck = deck;
		this.invalidate();
	}

	/** Re-render from the current model, discarding cached slide elements. */
	invalidate(): void {
		for (const el of this.slideCache.values()) this.options.editor?.detach(el);
		this.slideCache.clear();
		this.showSlide(this.index);
	}

	setDirty(dirty: boolean): void {
		this.saveButtonEl?.toggleClass("is-dirty", dirty);
	}

	destroy(): void {
		for (const el of this.slideCache.values()) this.options.editor?.detach(el);
		this.resizeObserver?.disconnect();
		this.thumbObserver?.disconnect();
		this.root.removeEventListener("keydown", this.onKeyDown);
		this.stageEl?.removeEventListener("wheel", this.onWheel);
		this.slideCache.clear();
		this.root.detach();
	}
}
