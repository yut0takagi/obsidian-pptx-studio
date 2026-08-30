/**
 * The text caret, drawn in screen space.
 *
 * A slide is scaled to fit its pane, and the browser's own caret is scaled with
 * it: at a two-thirds fit the caret is a sub-pixel hairline that is genuinely
 * hard to see against slide artwork. This one is painted outside the slide's
 * transform at a fixed width, so it stays the same weight at any zoom, with a
 * pale halo so it shows up on a dark background as readily as on a light one.
 *
 * The native caret is hidden with `caret-color` while this is up. Nothing else
 * about the editing changes — the selection, the caret's position and IME
 * composition all still belong to the platform; this only draws.
 */
export class CaretPainter {
	private el: HTMLElement | null = null;
	private box: HTMLElement | null = null;
	private frame = 0;

	start(box: HTMLElement): void {
		this.stop();
		this.box = box;
		box.style.caretColor = "transparent";

		const el = document.createElement("div");
		el.className = "pptx-caret";
		document.body.appendChild(el);
		this.el = el;

		document.addEventListener("selectionchange", this.schedule);
		// Capture, because the pane the slide sits in is what actually scrolls.
		window.addEventListener("scroll", this.schedule, true);
		window.addEventListener("resize", this.schedule);
		box.addEventListener("input", this.schedule);
		this.update();
	}

	stop(): void {
		if (this.frame) window.cancelAnimationFrame(this.frame);
		this.frame = 0;
		document.removeEventListener("selectionchange", this.schedule);
		window.removeEventListener("scroll", this.schedule, true);
		window.removeEventListener("resize", this.schedule);
		this.box?.removeEventListener("input", this.schedule);
		if (this.box) this.box.style.caretColor = "";
		this.box = null;
		this.el?.remove();
		this.el = null;
	}

	private schedule = (): void => {
		if (this.frame) return;
		this.frame = window.requestAnimationFrame(() => {
			this.frame = 0;
			this.update();
		});
	};

	private update(): void {
		const el = this.el;
		const box = this.box;
		if (!el || !box) return;
		const rect = this.caretRect(box);
		if (!rect) {
			el.style.display = "none";
			return;
		}
		Object.assign(el.style, {
			display: "block",
			left: `${rect.left}px`,
			top: `${rect.top}px`,
			height: `${rect.height}px`,
		});
		// Restart the blink, so the caret is solid the instant it moves rather
		// than invisible for the rest of the off half of the cycle.
		el.style.animation = "none";
		void el.offsetWidth;
		el.style.animation = "";
	}

	private caretRect(box: HTMLElement): { left: number; top: number; height: number } | null {
		const selection = window.getSelection();
		if (!selection || selection.rangeCount === 0) return null;
		const range = selection.getRangeAt(0);
		// A highlighted range draws itself; a caret on top of it says nothing.
		if (!range.collapsed) return null;
		if (!box.contains(range.startContainer)) return null;

		const rects = range.getClientRects();
		const rect = rects.length > 0 ? rects[rects.length - 1] : null;
		if (rect && rect.height > 0) return { left: rect.left, top: rect.top, height: rect.height };

		// An empty paragraph has no rect of its own, so the caret goes where the
		// first character would land: the line's own box, read the way it aligns.
		const node = range.startContainer;
		const host = node instanceof Element ? node : node.parentElement;
		const line = host?.getBoundingClientRect();
		if (!host || !line || line.height === 0) return null;
		const align = window.getComputedStyle(host).textAlign;
		const left =
			align === "center"
				? line.left + line.width / 2
				: align === "right" || align === "end"
					? line.right
					: line.left;
		return { left, top: line.top, height: line.height };
	}
}
