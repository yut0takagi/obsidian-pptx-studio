import { setIcon } from "obsidian";
import { t } from "../i18n";
import type { Deck } from "../pptx/types";
import { renderSlide } from "./renderSlide";
import {
	type PresenterAction,
	presentationOrder,
	presenterAction,
	startPosition,
	transitionPlan,
} from "./slideshow";

export interface PresenterOptions {
	deck: Deck;
	/** The slide the show starts from, counting from zero. */
	startIndex: number;
	/** The document to present in: the one the deck's leaf lives in. */
	doc: Document;
	/** Skip transitions; defaults to the system's reduced-motion setting. */
	reducedMotion?: boolean;
	/** Told which slide the show ended on, so the view can be left there. */
	onClose?: (slideIndex: number) => void;
}

/** How long the on-screen controls stay up after the pointer last moved. */
const CONTROLS_MS = 2500;

/**
 * A slide show over the whole screen.
 *
 * It asks for fullscreen, and where that is refused or does not exist — most
 * mobile builds — it covers the Obsidian window instead, which is as close as
 * a page can get. Slides are rendered the same way the deck view renders them,
 * scaled to fit, on black. Each slide's own transition is played on the way
 * in; going back shows the earlier slide straight away, as PowerPoint does.
 */
export class Presenter {
	readonly element: HTMLElement;
	private readonly options: PresenterOptions;
	private readonly doc: Document;
	private readonly win: Window;
	private readonly order: number[];
	private readonly frame: HTMLElement;
	private readonly controls: HTMLElement;
	private readonly counter: HTMLElement;
	private readonly reducedMotion: boolean;

	/** Position in `order`; `order.length` is the closing black screen. */
	private position: number;
	private scale = 1;
	private layer: HTMLElement | null = null;
	private running: { animations: Animation[]; leaving: HTMLElement | null } | null = null;
	private advanceTimer = 0;
	private controlsTimer = 0;
	private wasFullscreen = false;
	private closed = false;

	/** Open a show, or return null when every slide is hidden and there is nothing to show. */
	static open(options: PresenterOptions): Presenter | null {
		const order = presentationOrder(options.deck);
		const position = startPosition(order, options.startIndex);
		if (position === -1) return null;
		return new Presenter(options, order, position);
	}

	private constructor(options: PresenterOptions, order: number[], position: number) {
		this.options = options;
		this.doc = options.doc;
		this.win = options.doc.defaultView ?? window;
		this.order = order;
		this.position = position;
		this.reducedMotion =
			options.reducedMotion ??
			this.win.matchMedia?.("(prefers-reduced-motion: reduce)").matches ??
			false;

		this.element = this.doc.body.createDiv({ cls: "pptx-presenter is-windowed" });
		this.element.tabIndex = -1;
		this.element.setAttribute("role", "dialog");
		this.element.setAttribute("aria-label", options.deck.title);
		this.frame = this.element.createDiv({ cls: "pptx-presenter-frame" });

		this.controls = this.element.createDiv({ cls: "pptx-presenter-controls" });
		this.button("chevron-left", t("present.previous"), () => this.previous());
		this.counter = this.controls.createSpan({ cls: "pptx-presenter-counter" });
		this.button("chevron-right", t("present.next"), () => this.next());
		this.button("x", t("present.exit"), () => this.close());

		// Capturing on the window puts the show ahead of every other listener,
		// so Space or an arrow key cannot also reach an editor behind it.
		this.win.addEventListener("keydown", this.onKeyDown, true);
		this.win.addEventListener("resize", this.onResize);
		this.doc.addEventListener("fullscreenchange", this.onFullscreenChange);
		this.element.addEventListener("click", this.onClick);
		this.element.addEventListener("pointermove", this.onPointerMove);
		this.element.addEventListener("pointerdown", this.onPointerMove);

		this.fit();
		this.show(false);
		// The controls show for a moment at the start, so it is clear they exist.
		this.onPointerMove();
		this.element.focus();
		this.requestFullscreen();
	}

	/** The slide on screen, counting from zero, or null on the closing black screen. */
	get slideIndex(): number | null {
		return this.order[this.position] ?? null;
	}

	next(): void {
		if (this.position >= this.order.length) {
			this.close();
			return;
		}
		this.position += 1;
		this.show(true);
	}

	previous(): void {
		if (this.position === 0) return;
		this.position -= 1;
		this.show(false);
	}

	first(): void {
		this.position = 0;
		this.show(false);
	}

	last(): void {
		this.position = this.order.length - 1;
		this.show(false);
	}

	close(): void {
		if (this.closed) return;
		this.closed = true;
		this.win.clearTimeout(this.advanceTimer);
		this.win.clearTimeout(this.controlsTimer);
		this.settle();
		this.win.removeEventListener("keydown", this.onKeyDown, true);
		this.win.removeEventListener("resize", this.onResize);
		this.doc.removeEventListener("fullscreenchange", this.onFullscreenChange);
		if (this.doc.fullscreenElement === this.element) void this.doc.exitFullscreen?.().catch(() => undefined);
		this.element.detach();
		// The closing black screen is not a slide; the last one shown is.
		const index = this.order[Math.min(this.position, this.order.length - 1)];
		this.options.onClose?.(index);
	}

	// --------------------------------------------------------------- slides

	/** Put the current position on screen, playing its transition when `forward`. */
	private show(forward: boolean): void {
		this.win.clearTimeout(this.advanceTimer);
		// A press during a transition finishes it rather than stacking another.
		this.settle();

		const index = this.slideIndex;
		const incoming = this.frame.createDiv({ cls: "pptx-presenter-layer" });
		if (index === null) {
			incoming.addClass("pptx-presenter-end");
			incoming.createDiv({ cls: "pptx-presenter-end-text", text: t("present.end") });
		} else {
			const slideEl = renderSlide(this.options.deck, this.options.deck.slides[index]);
			slideEl.setCssStyles({ transform: `scale(${this.scale})` });
			incoming.appendChild(slideEl);
		}
		this.counter.setText(
			index === null ? "" : `${this.position + 1} / ${this.order.length}`,
		);

		const leaving = this.layer;
		this.layer = incoming;
		const transition = index === null ? null : this.options.deck.slides[index].transition;
		const plan = forward && leaving ? transitionPlan(transition, { reducedMotion: this.reducedMotion }) : null;

		if (!plan || !leaving || typeof incoming.animate !== "function") {
			leaving?.detach();
		} else {
			if (plan.outgoingOnTop) this.frame.insertBefore(incoming, leaving);
			const timing: KeyframeAnimationOptions = { duration: plan.duration, easing: plan.easing, fill: "both" };
			const animations: Animation[] = [];
			if (plan.incoming.length > 0) animations.push(incoming.animate(plan.incoming, timing));
			if (plan.outgoing) animations.push(leaving.animate(plan.outgoing, timing));
			const running = { animations, leaving };
			this.running = running;
			void Promise.all(animations.map((a) => a.finished))
				.catch(() => undefined)
				.then(() => {
					if (this.running === running) this.settle();
				});
		}

		const after = transition?.advanceAfter;
		if (index !== null && after !== null && after !== undefined) {
			this.advanceTimer = this.win.setTimeout(() => this.next(), (plan?.duration ?? 0) + after);
		}
	}

	/** End any transition in progress: the old slide goes, the new one stands as it is. */
	private settle(): void {
		const running = this.running;
		if (!running) return;
		this.running = null;
		for (const animation of running.animations) {
			// Fill "both" holds the end state; cancelling hands the layer back to its
			// own styles, which are that same end state.
			animation.cancel();
		}
		running.leaving?.detach();
	}

	/** Scale the slide to the largest size that fits the screen. */
	private fit(): void {
		const { width, height } = this.options.deck;
		const availableWidth = this.element.clientWidth || this.win.innerWidth || width;
		const availableHeight = this.element.clientHeight || this.win.innerHeight || height;
		this.scale = Math.max(0.05, Math.min(availableWidth / width, availableHeight / height));
		this.frame.setCssStyles({
			width: `${width * this.scale}px`,
			height: `${height * this.scale}px`,
		});
		for (const slideEl of Array.from(this.frame.querySelectorAll<HTMLElement>(".pptx-slide"))) {
			slideEl.setCssStyles({ transform: `scale(${this.scale})` });
		}
	}

	// --------------------------------------------------------------- chrome

	private button(icon: string, label: string, onClick: () => void): void {
		const btn = this.controls.createEl("button", { cls: "pptx-presenter-btn clickable-icon" });
		setIcon(btn, icon);
		btn.setAttribute("aria-label", label);
		btn.addEventListener("click", (event) => {
			event.stopPropagation();
			onClick();
		});
	}

	/** Ask for fullscreen; the window-covering fallback is already showing if it is refused. */
	private requestFullscreen(): void {
		if (typeof this.element.requestFullscreen !== "function") return;
		try {
			void this.element
				.requestFullscreen()
				.then(() => this.enteredFullscreen())
				.catch(() => undefined);
		} catch {
			// Some builds throw rather than reject; the fallback stands either way.
		}
	}

	private enteredFullscreen(): void {
		if (this.closed) return;
		this.wasFullscreen = true;
		this.element.removeClass("is-windowed");
		this.fit();
	}

	// --------------------------------------------------------------- events

	private onKeyDown = (event: KeyboardEvent): void => {
		const action = presenterAction(event);
		if (!action) return;
		event.preventDefault();
		event.stopPropagation();
		this.perform(action);
	};

	private perform(action: PresenterAction): void {
		switch (action) {
			case "next":
				this.next();
				break;
			case "previous":
				this.previous();
				break;
			case "first":
				this.first();
				break;
			case "last":
				this.last();
				break;
			case "exit":
				this.close();
				break;
		}
	}

	/**
	 * A click advances, as it does in PowerPoint. On a touch screen there are no
	 * keys to go back with, so a tap on the left third of the screen does that.
	 */
	private onClick = (event: MouseEvent): void => {
		if (this.controls.contains(event.target as Node)) return;
		const touch = (event as PointerEvent).pointerType === "touch";
		const width = this.element.clientWidth;
		if (touch && width > 0 && event.clientX < width / 3) this.previous();
		else this.next();
	};

	private onPointerMove = (): void => {
		this.element.addClass("is-showing-controls");
		this.win.clearTimeout(this.controlsTimer);
		this.controlsTimer = this.win.setTimeout(
			() => this.element.removeClass("is-showing-controls"),
			CONTROLS_MS,
		);
	};

	private onResize = (): void => this.fit();

	/**
	 * Escape in fullscreen is the browser's: it leaves fullscreen, and the page
	 * only hears about it afterwards. Leaving fullscreen by any route ends the
	 * show, rather than stranding it in a window-sized overlay.
	 */
	private onFullscreenChange = (): void => {
		if (this.doc.fullscreenElement === this.element) {
			this.enteredFullscreen();
			return;
		}
		if (this.wasFullscreen) this.close();
	};
}
