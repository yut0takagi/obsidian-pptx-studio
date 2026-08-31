export interface PaneSplitterOptions {
	/** The pane above the divider: its height is what a drag changes. */
	above: HTMLElement;
	/** Starting height in pixels, and where a finished drag reports the new one. */
	height: number;
	onHeight: (height: number) => void;
	/** Height the divider snaps back to on a double-click. */
	defaultHeight: number;
}

const MIN = 60;
/** Leave at least this much for the pane underneath. */
const KEEP_BELOW = 80;

/**
 * The draggable line between two stacked panes.
 *
 * It only ever sets the height of the pane above; the one below takes whatever
 * is left, so the pair always fills the column exactly and neither can be
 * pushed off the bottom.
 */
export class PaneSplitter {
	private readonly el: HTMLElement;
	private drag: { pointerId: number; startY: number; startHeight: number } | null = null;

	constructor(
		containerEl: HTMLElement,
		private readonly options: PaneSplitterOptions,
	) {
		this.el = containerEl.createDiv({ cls: "pptx-pane-splitter" });
		this.el.addEventListener("pointerdown", this.onPointerDown);
		this.el.addEventListener("dblclick", () => this.apply(this.options.defaultHeight));
		this.apply(options.height);
	}

	/** Hidden while one of the two panes is collapsed: there is nothing to split. */
	setEnabled(enabled: boolean): void {
		this.el.toggleClass("is-hidden", !enabled);
	}

	setHeight(height: number): void {
		this.apply(height);
	}

	private apply(height: number): void {
		const container = this.el.parentElement;
		const room = container ? container.getBoundingClientRect().height : 0;
		const max = room > 0 ? Math.max(MIN, room - KEEP_BELOW) : Number.MAX_SAFE_INTEGER;
		const clamped = Math.round(Math.max(MIN, Math.min(max, height)));
		this.options.above.setCssStyles({ height: `${clamped}px` });
		this.options.onHeight(clamped);
	}

	private onPointerDown = (event: PointerEvent): void => {
		if (event.button !== 0) return;
		event.preventDefault();
		this.drag = {
			pointerId: event.pointerId,
			startY: event.clientY,
			startHeight: this.options.above.getBoundingClientRect().height,
		};
		this.el.addClass("is-dragging");
		window.addEventListener("pointermove", this.onPointerMove);
		window.addEventListener("pointerup", this.onPointerUp);
		window.addEventListener("pointercancel", this.onPointerUp);
	};

	private onPointerMove = (event: PointerEvent): void => {
		const drag = this.drag;
		if (!drag || event.pointerId !== drag.pointerId) return;
		this.apply(drag.startHeight + (event.clientY - drag.startY));
	};

	private onPointerUp = (event: PointerEvent): void => {
		if (!this.drag || event.pointerId !== this.drag.pointerId) return;
		this.drag = null;
		this.el.removeClass("is-dragging");
		this.unlisten();
	};

	private unlisten(): void {
		window.removeEventListener("pointermove", this.onPointerMove);
		window.removeEventListener("pointerup", this.onPointerUp);
		window.removeEventListener("pointercancel", this.onPointerUp);
	}

	destroy(): void {
		this.unlisten();
		this.el.detach();
	}
}
