import type { SlideTransition } from "../pptx/transition";
import type { Deck } from "../pptx/types";

/**
 * The rules of presenting, as data: which slides a show runs through, where it
 * starts, what a key does, and what a transition is as keyframes. The DOM that
 * carries them out lives in Presenter.ts; nothing here touches it, so each
 * rule can be checked without an animation to watch.
 */

/** Indices into `deck.slides` of the slides a show runs through, in order. */
export function presentationOrder(deck: Deck): number[] {
	const out: number[] = [];
	deck.slides.forEach((slide, i) => {
		if (!slide.hidden) out.push(i);
	});
	return out;
}

/**
 * Where in `order` a show started from slide `from` begins.
 *
 * Starting on a hidden slide begins at the next one that is shown, which is
 * where pressing "next" would have gone; with nothing shown after it, the show
 * starts on the last slide that is. -1 means there is nothing to show.
 */
export function startPosition(order: number[], from: number): number {
	if (order.length === 0) return -1;
	const at = order.findIndex((index) => index >= from);
	return at === -1 ? order.length - 1 : at;
}

export type PresenterAction = "next" | "previous" | "first" | "last" | "exit";

const ACTION_BY_KEY: Record<string, PresenterAction> = {
	ArrowRight: "next",
	ArrowDown: "next",
	PageDown: "next",
	" ": "next",
	Enter: "next",
	n: "next",
	N: "next",
	ArrowLeft: "previous",
	ArrowUp: "previous",
	PageUp: "previous",
	Backspace: "previous",
	p: "previous",
	P: "previous",
	Home: "first",
	End: "last",
	Escape: "exit",
};

/**
 * What a key press means during a show. Presentation clickers send PageUp and
 * PageDown; keyboards send the arrows and Space. A key held with a modifier is
 * a shortcut, and is left for Obsidian.
 */
export function presenterAction(event: {
	key: string;
	altKey?: boolean;
	ctrlKey?: boolean;
	metaKey?: boolean;
}): PresenterAction | null {
	if (event.altKey || event.ctrlKey || event.metaKey) return null;
	return ACTION_BY_KEY[event.key] ?? null;
}

export interface TransitionPlan {
	duration: number;
	easing: string;
	/** Keyframes for the slide coming in; empty when it simply sits underneath. */
	incoming: Keyframe[];
	/** Keyframes for the slide going out, or null when it is just covered. */
	outgoing: Keyframe[] | null;
	/** Draw the outgoing slide above the incoming one, for effects that take it away. */
	outgoingOnTop: boolean;
}

/** Unit offsets for the way a slide travels: "l" moves it left. */
const TRAVEL: Record<string, [number, number]> = {
	l: [-1, 0],
	r: [1, 0],
	u: [0, -1],
	d: [0, 1],
	lu: [-1, -1],
	ru: [1, -1],
	ld: [-1, 1],
	rd: [1, 1],
};

const translate = (x: number, y: number): string => `translate(${x * 100 || 0}%, ${y * 100 || 0}%)`;
const inset = (top: number, right: number, bottom: number, left: number): string =>
	`inset(${top}% ${right}% ${bottom}% ${left}%)`;
const OPEN = inset(0, 0, 0, 0);

/**
 * The keyframes a transition is played with, or null to switch instantly.
 *
 * A transition belongs to the slide it brings in. `dir` is the way the slides
 * travel, so a push to the left brings the new slide in from the right. Motion
 * is skipped altogether for anyone who has asked their system for less of it.
 */
export function transitionPlan(
	transition: SlideTransition | null,
	options: { reducedMotion: boolean },
): TransitionPlan | null {
	if (!transition || options.reducedMotion) return null;
	if (transition.effect === "cut" || transition.duration <= 0) return null;

	const plan = (patch: Partial<TransitionPlan>): TransitionPlan => ({
		duration: transition.duration,
		easing: "ease-in-out",
		incoming: [],
		outgoing: null,
		outgoingOnTop: false,
		...patch,
	});
	const [dx, dy] = TRAVEL[transition.direction ?? "l"] ?? TRAVEL.l;

	switch (transition.effect) {
		case "push":
			return plan({
				incoming: [{ transform: translate(-dx, -dy) }, { transform: translate(0, 0) }],
				outgoing: [{ transform: translate(0, 0) }, { transform: translate(dx, dy) }],
			});
		case "cover":
			return plan({
				incoming: [{ transform: translate(-dx, -dy) }, { transform: translate(0, 0) }],
			});
		case "pull":
			return plan({
				outgoing: [{ transform: translate(0, 0) }, { transform: translate(dx, dy) }],
				outgoingOnTop: true,
			});
		case "wipe": {
			// The edge the reveal starts from is the one the wipe travels away from.
			const from =
				dx < 0 ? inset(0, 0, 0, 100) : dx > 0 ? inset(0, 100, 0, 0) : dy < 0 ? inset(100, 0, 0, 0) : inset(0, 0, 100, 0);
			return plan({ incoming: [{ clipPath: from }, { clipPath: OPEN }] });
		}
		case "split": {
			const closed = transition.orientation === "vert" ? inset(0, 50, 0, 50) : inset(50, 0, 50, 0);
			return transition.direction === "in"
				? plan({ outgoing: [{ clipPath: OPEN }, { clipPath: closed }], outgoingOnTop: true })
				: plan({ incoming: [{ clipPath: closed }, { clipPath: OPEN }] });
		}
		case "zoom":
			return transition.direction === "in"
				? plan({
						incoming: [
							{ transform: "scale(0.3)", opacity: 0 },
							{ transform: "scale(1)", opacity: 1 },
						],
					})
				: plan({
						outgoing: [
							{ transform: "scale(1)", opacity: 1 },
							{ transform: "scale(1.5)", opacity: 0 },
						],
						outgoingOnTop: true,
					});
		case "circle":
			return plan({
				incoming: [{ clipPath: "circle(0% at 50% 50%)" }, { clipPath: "circle(75% at 50% 50%)" }],
			});
		case "fade":
		default:
			if (transition.throughBlack) {
				// The stage behind the slides is black, so fading the old slide out
				// and only then the new one in passes through it.
				return plan({
					easing: "linear",
					outgoing: [{ opacity: 1 }, { opacity: 0, offset: 0.5 }, { opacity: 0 }],
					incoming: [{ opacity: 0 }, { opacity: 0, offset: 0.5 }, { opacity: 1 }],
				});
			}
			return plan({ easing: "linear", incoming: [{ opacity: 0 }, { opacity: 1 }] });
	}
}
