import { attr, boolAttr, child, numAttr } from "./xml";

/**
 * A slide's transition, as presenting needs it.
 *
 * The editor reads transitions too (src/ooxml/animation.ts), but only the ones
 * it can write back, and only from a bare `p:transition`. Playing a deck has to
 * take whatever the deck carries. A current PowerPoint wraps most transitions
 * in `mc:AlternateContent`: a `p14` choice holding the real duration and, for
 * the 2010-and-later effects, an effect older readers do not know, next to a
 * fallback in plain PresentationML. Both halves are read here, and an effect
 * with no drawing of its own still plays — as a fade — rather than as nothing.
 */

/** The effects there is a drawing for; anything else is played as a fade. */
export type TransitionEffect =
	| "cut"
	| "fade"
	| "push"
	| "wipe"
	| "cover"
	| "pull"
	| "split"
	| "zoom"
	| "circle";

export interface SlideTransition {
	effect: TransitionEffect;
	/** Milliseconds. Zero for a cut. */
	duration: number;
	/**
	 * Which way it travels: "l" / "r" / "u" / "d" (and the diagonals, for cover
	 * and pull) is the direction the slides move in; "in" / "out" for split and
	 * zoom. Null for effects that do not travel.
	 */
	direction: string | null;
	/** Split only: whether the opening runs along a horizontal or a vertical line. */
	orientation: "horz" | "vert" | null;
	/** A fade or cut that passes through black on the way. */
	throughBlack: boolean;
	/** Advance on its own this many milliseconds after the slide appears; null waits for the presenter. */
	advanceAfter: number | null;
}

/**
 * How long the named speeds last. PresentationML leaves the figures to the
 * application; these are the ones PowerPoint plays them at.
 */
const SPEED_MS: Record<string, number> = { fast: 500, med: 750, slow: 1000 };

const TRAVELLING = new Set(["push", "wipe", "cover", "pull"]);

/** Effects with a drawing, by element name, and the ones close enough to borrow one. */
const EFFECT_BY_ELEMENT: Record<string, TransitionEffect> = {
	cut: "cut",
	fade: "fade",
	push: "push",
	wipe: "wipe",
	cover: "cover",
	pull: "pull",
	split: "split",
	zoom: "zoom",
	circle: "circle",
};

/** PresentationML 2006 transitions that have no drawing of their own and play as a fade. */
const PLAIN_ELEMENTS = new Set([
	"blinds",
	"checker",
	"comb",
	"diamond",
	"dissolve",
	"newsflash",
	"plus",
	"random",
	"randomBar",
	"strips",
	"wedge",
	"wheel",
]);

const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";

/** The transition on a slide, or null when it has none. */
export function parseSlideTransition(slide: Element | null): SlideTransition | null {
	if (!slide) return null;
	const direct = child(slide, "transition");
	if (direct) return readTransition(direct, null);

	for (let n = slide.firstElementChild; n; n = n.nextElementSibling) {
		if (n.localName !== "AlternateContent") continue;
		const choice = child(child(n, "Choice"), "transition");
		const fallback = child(child(n, "Fallback"), "transition");
		if (!choice && !fallback) continue;
		// The choice knows the real duration; the fallback knows an effect that
		// is drawable when the choice's is one of the newer ones.
		if (choice && effectElement(choice, true)) return readTransition(choice, null);
		return readTransition(fallback ?? choice!, choice);
	}
	return null;
}

/**
 * The element naming the effect. With `known` set, only effects there is a
 * drawing for count, which is how a p14-only choice is told apart from one
 * that merely carries a p14 duration.
 */
function effectElement(transition: Element, known: boolean): Element | null {
	for (let n = transition.firstElementChild; n; n = n.nextElementSibling) {
		// Sounds and extension lists sit alongside the effect.
		if (n.localName === "sndAc" || n.localName === "extLst") continue;
		if (!known) return n;
		if (n.namespaceURI === P_NS && (EFFECT_BY_ELEMENT[n.localName] || PLAIN_ELEMENTS.has(n.localName))) {
			return n;
		}
	}
	return null;
}

function readTransition(transition: Element, timingFrom: Element | null): SlideTransition {
	const effectEl = effectElement(transition, false);
	const name = effectEl?.localName ?? null;
	const effect: TransitionEffect =
		name === null ? "cut" : (EFFECT_BY_ELEMENT[name] ?? "fade");

	const timing = timingFrom ?? transition;
	const explicit = numAttr(timing, "dur");
	const named = SPEED_MS[attr(timing, "spd") ?? attr(transition, "spd") ?? "med"] ?? SPEED_MS.med;
	const duration = effect === "cut" ? 0 : Math.max(0, explicit ?? named);

	let direction: string | null = null;
	let orientation: SlideTransition["orientation"] = null;
	if (TRAVELLING.has(effect)) {
		direction = attr(effectEl, "dir") ?? "l";
	} else if (effect === "split") {
		direction = attr(effectEl, "dir") === "in" ? "in" : "out";
		orientation = attr(effectEl, "orient") === "vert" ? "vert" : "horz";
	} else if (effect === "zoom") {
		direction = attr(effectEl, "dir") === "in" ? "in" : "out";
	}

	const advance = numAttr(transition, "advTm");

	return {
		effect,
		duration,
		direction,
		orientation,
		throughBlack: boolAttr(effectEl, "thruBlk") ?? false,
		advanceAfter: advance !== null && advance >= 0 ? advance : null,
	};
}
