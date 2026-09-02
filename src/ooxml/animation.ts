/**
 * Slide transitions and the animation main sequence.
 *
 * `p:timing` is the most involved corner of PresentationML: a tree of nested
 * time nodes where an entrance effect on one shape is four levels of `p:par`
 * deep, and where every `p:cTn` carries an id that has to be unique across the
 * slide. PowerPoint is unforgiving about the shape of it — a tree that is
 * merely plausible makes the deck refuse to open — so the structure written
 * here is the one PowerPoint itself emits, rather than the smallest one the
 * schema would allow.
 *
 * Two consequences follow, and both are deliberate. Ids are renumbered from
 * scratch after every change, because keeping them consistent incrementally is
 * exactly the kind of bookkeeping that silently drifts. And a timing tree this
 * code did not write is read but never edited in place: what it can say about
 * such a slide is what is there, and the only change it offers is removal.
 */
import { attr, child, children } from "../pptx/xml";
import { p as pEl } from "./tree";

// ------------------------------------------------------------- transitions

export type TransitionKind =
	| "fade"
	| "cut"
	| "push"
	| "wipe"
	| "cover"
	| "dissolve"
	| "split"
	| "randomBar";

export type TransitionSpeed = "slow" | "med" | "fast";

/** The transitions that take a direction, and the directions they take. */
export const DIRECTED: Record<string, string[]> = {
	push: ["l", "r", "u", "d"],
	wipe: ["l", "r", "u", "d"],
	cover: ["l", "r", "u", "d"],
};

export interface Transition {
	kind: TransitionKind;
	speed: TransitionSpeed;
	/** "l" / "r" / "u" / "d" for the transitions that travel; null otherwise. */
	direction: string | null;
}

const TRANSITION_KINDS = new Set<string>([
	"fade",
	"cut",
	"push",
	"wipe",
	"cover",
	"dissolve",
	"split",
	"randomBar",
]);

export function readTransition(slide: Element): Transition | null {
	const el = child(slide, "transition");
	if (!el) return null;
	for (let n = el.firstElementChild; n; n = n.nextElementSibling) {
		if (!TRANSITION_KINDS.has(n.localName)) continue;
		const speed = attr(el, "spd");
		return {
			kind: n.localName as TransitionKind,
			speed: speed === "slow" || speed === "fast" ? speed : "med",
			direction: attr(n, "dir"),
		};
	}
	return null;
}

/**
 * Set, change or clear the slide's transition.
 *
 * `p:transition` has one place it may sit — after `p:clrMapOvr` and before
 * `p:timing` — and a slide whose children are out of schema order is a slide
 * PowerPoint rejects, so the position is computed rather than appended to.
 */
export function writeTransition(slide: Element, value: Transition | null): boolean {
	const existing = child(slide, "transition");
	if (!value) {
		if (!existing) return false;
		slide.removeChild(existing);
		return true;
	}

	const doc = slide.ownerDocument;
	const el = pEl(doc, "transition");
	el.setAttribute("spd", value.speed);
	const inner = pEl(doc, value.kind);
	const directions = DIRECTED[value.kind];
	if (directions) inner.setAttribute("dir", directions.includes(value.direction ?? "") ? value.direction! : directions[0]);
	el.appendChild(inner);

	if (existing) slide.replaceChild(el, existing);
	else slide.insertBefore(el, child(slide, "timing"));
	return true;
}

// -------------------------------------------------------------- animations

export type EffectKind = "fade" | "wipe" | "flyIn";
export type Trigger = "click" | "withPrev" | "afterPrev";

/** presetID is how PowerPoint names an effect; these are the three offered. */
const PRESET_ID: Record<EffectKind, string> = { fade: "10", wipe: "22", flyIn: "2" };
const EFFECT_BY_PRESET: Record<string, EffectKind> = { "10": "fade", "22": "wipe", "2": "flyIn" };

const NODE_TYPE: Record<Trigger, string> = {
	click: "clickEffect",
	withPrev: "withEffect",
	afterPrev: "afterEffect",
};
const TRIGGER_BY_NODE_TYPE: Record<string, Trigger> = {
	clickEffect: "click",
	withEffect: "withPrev",
	afterEffect: "afterPrev",
};

/** How long an entrance runs, in milliseconds. */
const DURATION = "500";

export interface AnimationEntry {
	/** The p:cNvPr id of the shape the effect targets. */
	shapeId: string;
	/** "other" for an effect written by PowerPoint that this code does not model. */
	effect: EffectKind | "other";
	trigger: Trigger;
	/** Position in the main sequence, counting from zero. */
	index: number;
}

function el(
	doc: Document,
	name: string,
	attrs: Record<string, string> = {},
	kids: Element[] = [],
): Element {
	const node = pEl(doc, name);
	for (const [key, value] of Object.entries(attrs)) node.setAttribute(key, value);
	for (const kid of kids) node.appendChild(kid);
	return node;
}

/** The main sequence's child list, which holds one entry per click group. */
function mainSeqList(slide: Element): Element | null {
	const timing = child(slide, "timing");
	const tnLst = child(timing, "tnLst");
	const root = child(tnLst, "par");
	const rootTn = child(root, "cTn");
	for (const node of children(child(rootTn, "childTnLst"), "seq")) {
		if (attr(node, "nodeType") === "mainSeq" || attr(child(node, "cTn"), "nodeType") === "mainSeq") {
			return child(child(node, "cTn"), "childTnLst");
		}
	}
	return null;
}

/** Every effect in the main sequence, in the order it plays. */
function effectPars(slide: Element): Element[] {
	const list = mainSeqList(slide);
	if (!list) return [];
	const out: Element[] = [];
	for (const clickGroup of children(list, "par")) {
		const groupList = child(child(clickGroup, "cTn"), "childTnLst");
		for (const middle of children(groupList, "par")) {
			const middleList = child(child(middle, "cTn"), "childTnLst");
			for (const effect of children(middleList, "par")) out.push(effect);
		}
	}
	return out;
}

function spidOf(effect: Element): string | null {
	const target = effect.getElementsByTagName("*");
	for (let i = 0; i < target.length; i++) {
		if (target[i].localName === "spTgt") return attr(target[i], "spid");
	}
	return null;
}

export function readAnimations(slide: Element): AnimationEntry[] {
	const out: AnimationEntry[] = [];
	for (const effect of effectPars(slide)) {
		const cTn = child(effect, "cTn");
		const shapeId = spidOf(effect);
		if (!cTn || !shapeId) continue;
		const preset = attr(cTn, "presetID") ?? "";
		out.push({
			shapeId,
			effect: EFFECT_BY_PRESET[preset] ?? "other",
			trigger: TRIGGER_BY_NODE_TYPE[attr(cTn, "nodeType") ?? ""] ?? "click",
			// Numbered after the skips, so it always indexes back into this list.
			index: out.length,
		});
	}
	return out;
}

// ------------------------------------------------------------ writing timing

function condList(doc: Document, name: string, delay: string): Element {
	return el(doc, name, {}, [el(doc, "cond", { delay })]);
}

function slideCond(doc: Document, name: string, event: string): Element {
	return el(doc, name, {}, [
		el(doc, "cond", { evt: event, delay: "0" }, [el(doc, "tgtEl", {}, [el(doc, "sldTgt")])]),
	]);
}

function target(doc: Document, shapeId: string): Element {
	return el(doc, "tgtEl", {}, [el(doc, "spTgt", { spid: shapeId })]);
}

/** The behaviour every entrance opens with: make the shape visible. */
function showBehaviour(doc: Document, shapeId: string): Element {
	return el(doc, "set", {}, [
		el(doc, "cBhvr", {}, [
			el(doc, "cTn", { dur: "1", fill: "hold" }, [condList(doc, "stCondLst", "0")]),
			target(doc, shapeId),
			el(doc, "attrNameLst", {}, [attrName(doc, "style.visibility")]),
		]),
		el(doc, "to", {}, [el(doc, "strVal", { val: "visible" })]),
	]);
}

function attrName(doc: Document, name: string): Element {
	const node = pEl(doc, "attrName");
	node.textContent = name;
	return node;
}

/** One axis of a fly-in: hold the value, then run it to where the shape sits. */
function slide1d(doc: Document, shapeId: string, axis: "ppt_x" | "ppt_y", from: string): Element {
	return el(doc, "anim", { calcmode: "lin", valueType: "num" }, [
		el(doc, "cBhvr", { additive: "base" }, [
			el(doc, "cTn", { dur: DURATION, fill: "hold" }),
			target(doc, shapeId),
			el(doc, "attrNameLst", {}, [attrName(doc, axis)]),
		]),
		el(doc, "tavLst", {}, [
			el(doc, "tav", { tm: "0" }, [el(doc, "val", {}, [el(doc, "strVal", { val: from })])]),
			el(doc, "tav", { tm: "100000" }, [
				el(doc, "val", {}, [el(doc, "strVal", { val: `#${axis}` })]),
			]),
		]),
	]);
}

function effectBody(doc: Document, effect: EffectKind, shapeId: string): Element[] {
	const show = showBehaviour(doc, shapeId);
	switch (effect) {
		case "fade":
			return [
				show,
				el(doc, "animEffect", { transition: "in", filter: "fade" }, [
					el(doc, "cBhvr", {}, [el(doc, "cTn", { dur: DURATION }), target(doc, shapeId)]),
				]),
			];
		case "wipe":
			return [
				show,
				el(doc, "animEffect", { transition: "in", filter: "wipe(up)" }, [
					el(doc, "cBhvr", {}, [el(doc, "cTn", { dur: DURATION }), target(doc, shapeId)]),
				]),
			];
		case "flyIn":
			return [
				show,
				slide1d(doc, shapeId, "ppt_x", "#ppt_x"),
				slide1d(doc, shapeId, "ppt_y", "1+#ppt_h/2"),
			];
	}
}

/** The innermost par: one effect on one shape. */
function buildEffect(
	doc: Document,
	shapeId: string,
	effect: EffectKind,
	trigger: Trigger,
): Element {
	return el(doc, "par", {}, [
		el(
			doc,
			"cTn",
			{
				presetID: PRESET_ID[effect],
				presetClass: "entr",
				presetSubtype: effect === "fade" ? "0" : "4",
				fill: "hold",
				grpId: "0",
				nodeType: NODE_TYPE[trigger],
			},
			[condList(doc, "stCondLst", "0"), el(doc, "childTnLst", {}, effectBody(doc, effect, shapeId))],
		),
	]);
}

/** The middle par, which is what an "after previous" effect waits on. */
function buildMiddle(doc: Document, effect: Element): Element {
	return el(doc, "par", {}, [
		el(doc, "cTn", { fill: "hold" }, [
			condList(doc, "stCondLst", "0"),
			el(doc, "childTnLst", {}, [effect]),
		]),
	]);
}

/** The outer par: everything that runs off one click. */
function buildClickGroup(doc: Document, middle: Element): Element {
	return el(doc, "par", {}, [
		el(doc, "cTn", { fill: "hold" }, [
			condList(doc, "stCondLst", "indefinite"),
			el(doc, "childTnLst", {}, [middle]),
		]),
	]);
}

/** Create the timing tree a slide with no animations has never had. */
function ensureTiming(slide: Element): Element {
	const existing = mainSeqList(slide);
	if (existing) return existing;

	const doc = slide.ownerDocument;
	const seqList = el(doc, "childTnLst");
	const seq = el(doc, "seq", { concurrent: "1", nextAc: "seek" }, [
		el(doc, "cTn", { id: "2", dur: "indefinite", nodeType: "mainSeq" }, [seqList]),
		slideCond(doc, "prevCondLst", "onPrev"),
		slideCond(doc, "nextCondLst", "onNext"),
	]);
	const timing = el(doc, "timing", {}, [
		el(doc, "tnLst", {}, [
			el(doc, "par", {}, [
				el(
					doc,
					"cTn",
					{ id: "1", dur: "indefinite", restart: "never", nodeType: "tmRoot" },
					[el(doc, "childTnLst", {}, [seq])],
				),
			]),
		]),
	]);

	const old = child(slide, "timing");
	if (old) slide.replaceChild(timing, old);
	else slide.appendChild(timing);
	return seqList;
}

/**
 * Give every time node a fresh id.
 *
 * Ids only have to be unique within the slide, and nothing refers to them from
 * outside the tree, so renumbering in document order after each change is both
 * safe and the only version of this that cannot drift.
 */
function renumber(slide: Element): void {
	const timing = child(slide, "timing");
	if (!timing) return;
	let next = 1;
	const walk = (node: Element): void => {
		if (node.localName === "cTn") node.setAttribute("id", String(next++));
		for (let n = node.firstElementChild; n; n = n.nextElementSibling) walk(n);
	};
	walk(timing);
}

export function addAnimation(
	slide: Element,
	shapeId: string,
	effect: EffectKind,
	trigger: Trigger,
): boolean {
	const list = ensureTiming(slide);
	const doc = slide.ownerDocument;
	const groups = children(list, "par");

	// "With previous" joins the effect already running; "after previous" waits
	// for it but stays on the same click. Both need something to be previous to,
	// so the first effect on a slide is always a click however it was asked for
	// — writing "with previous" there would produce a node that says it runs
	// alongside something and a group that says it waits for a click.
	const last = groups[groups.length - 1] ?? null;
	const effective: Trigger = last ? trigger : "click";
	const effectPar = buildEffect(doc, shapeId, effect, effective);

	if (effective === "click" || !last) {
		list.appendChild(buildClickGroup(doc, buildMiddle(doc, effectPar)));
	} else {
		const groupList = child(child(last, "cTn"), "childTnLst");
		if (!groupList) return false;
		if (effective === "withPrev") {
			const middles = children(groupList, "par");
			const middleList = child(child(middles[middles.length - 1], "cTn"), "childTnLst");
			if (!middleList) return false;
			middleList.appendChild(effectPar);
		} else {
			groupList.appendChild(buildMiddle(doc, effectPar));
		}
	}
	renumber(slide);
	return true;
}

/** Take out an effect, and any wrapper it leaves empty behind it. */
function pruneUpwards(node: Element): void {
	let current: Element | null = node;
	while (current) {
		const parent: Element | null = current.parentElement;
		current.remove();
		if (!parent) return;
		// A childTnLst with nothing in it, and the cTn and par around it, are
		// meaningless on their own — and a p:par with no children is invalid.
		if (parent.localName !== "childTnLst" || parent.firstElementChild !== null) return;
		const cTn: Element | null = parent.parentElement;
		if (!cTn || cTn.localName !== "cTn") return;
		const par: Element | null = cTn.parentElement;
		if (!par || par.localName !== "par") return;
		current = par;
	}
}

export function removeAnimationsFor(slide: Element, shapeId: string): number {
	const targets = effectPars(slide).filter((par) => spidOf(par) === shapeId);
	for (const par of targets) pruneUpwards(par);
	if (targets.length > 0) tidy(slide);
	return targets.length;
}

export function removeAnimationAt(slide: Element, index: number): boolean {
	const par = effectPars(slide)[index];
	if (!par) return false;
	pruneUpwards(par);
	tidy(slide);
	return true;
}

export function clearAnimations(slide: Element): boolean {
	const timing = child(slide, "timing");
	if (!timing) return false;
	slide.removeChild(timing);
	return true;
}

/** Drop a main sequence that has been emptied, then renumber what is left. */
function tidy(slide: Element): void {
	if (effectPars(slide).length === 0) {
		clearAnimations(slide);
		return;
	}
	renumber(slide);
}

/**
 * Move an effect to another position in the sequence.
 *
 * The effect is rebuilt at its destination rather than moved, because where it
 * lands decides how it is wrapped: an effect that becomes the first of a click
 * group needs the group around it that its old position already had.
 */
export function moveAnimation(slide: Element, from: number, to: number): boolean {
	const entries = readAnimations(slide);
	if (from === to || from < 0 || to < 0 || from >= entries.length || to >= entries.length) {
		return false;
	}
	const reordered = [...entries];
	const [moved] = reordered.splice(from, 1);
	reordered.splice(to, 0, moved);
	return rewrite(slide, reordered);
}

/** Rebuild the whole main sequence from a list of effects. */
function rewrite(slide: Element, entries: AnimationEntry[]): boolean {
	// Anything this code cannot express would be lost by a rebuild, so a slide
	// carrying one is left exactly as it is.
	if (entries.some((entry) => entry.effect === "other")) return false;
	clearAnimations(slide);
	for (const entry of entries) {
		addAnimation(slide, entry.shapeId, entry.effect as EffectKind, entry.trigger);
	}
	return true;
}

/** True when every effect on the slide is one this code could rebuild. */
export function canReorder(slide: Element): boolean {
	const entries = readAnimations(slide);
	return entries.length > 1 && entries.every((entry) => entry.effect !== "other");
}
