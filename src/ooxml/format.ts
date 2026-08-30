import { PT_TO_PX } from "../pptx/types";
import { child, children } from "../pptx/xml";
import { A_NS } from "./tree";

/**
 * Character and paragraph formatting, written straight into the run and
 * paragraph properties the deck already has.
 *
 * OOXML is order-sensitive: a:rPr's children must appear in schema order or
 * PowerPoint rejects the file. Everything here goes through `setOrderedChild`
 * so a new element lands in the right place rather than wherever it was easiest
 * to append.
 */

const RPR_ORDER = [
	"ln",
	"noFill",
	"solidFill",
	"gradFill",
	"blipFill",
	"pattFill",
	"grpFill",
	"effectLst",
	"effectDag",
	"highlight",
	"uLnTx",
	"uLn",
	"uFillTx",
	"uFill",
	"latin",
	"ea",
	"cs",
	"sym",
	"hlinkClick",
	"hlinkMouseOver",
	"rtl",
	"extLst",
];

const PPR_ORDER = [
	"lnSpc",
	"spcBef",
	"spcAft",
	"buClrTx",
	"buClr",
	"buSzTx",
	"buSzPct",
	"buSzPts",
	"buFontTx",
	"buFont",
	"buNone",
	"buAutoNum",
	"buChar",
	"tabLst",
	"defRPr",
	"extLst",
];

export const SPPR_ORDER = [
	"xfrm",
	"custGeom",
	"prstGeom",
	"noFill",
	"solidFill",
	"gradFill",
	"blipFill",
	"pattFill",
	"grpFill",
	"ln",
	"effectLst",
	"effectDag",
	"scene3d",
	"sp3d",
	"extLst",
];

/**
 * Put `node` where the schema expects it among its siblings, replacing any
 * element of the same name. Passing null removes it.
 */
export function setOrderedChild(
	parent: Element,
	name: string,
	order: string[],
	node: Element | null,
): void {
	const existing = child(parent, name);
	if (existing) parent.removeChild(existing);
	if (!node) return;

	const rank = order.indexOf(name);
	let anchor: Element | null = null;
	for (let n = parent.firstElementChild; n; n = n.nextElementSibling) {
		const otherRank = order.indexOf(n.localName);
		if (otherRank === -1) continue;
		if (otherRank > rank) {
			anchor = n;
			break;
		}
	}
	parent.insertBefore(node, anchor);
}

function solidFillElement(doc: Document, color: string): Element {
	const fill = doc.createElementNS(A_NS, "a:solidFill");
	const clr = doc.createElementNS(A_NS, "a:srgbClr");
	clr.setAttribute("val", color.replace(/^#/, "").toUpperCase());
	fill.appendChild(clr);
	return fill;
}

// ------------------------------------------------------------------ runs

export interface RunPatch {
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	strike?: boolean;
	/** Points. */
	size?: number;
	/** "#rrggbb", or null to drop an explicit colour and inherit again. */
	color?: string | null;
	font?: string | null;
}

/** Get or create the a:rPr of an a:r / a:fld, which must be its first child. */
export function runProps(runEl: Element): Element {
	const existing = child(runEl, "rPr");
	if (existing) return existing;
	const rPr = runEl.ownerDocument.createElementNS(A_NS, "a:rPr");
	rPr.setAttribute("lang", "en-US");
	runEl.insertBefore(rPr, runEl.firstChild);
	return rPr;
}

export function applyRunPatch(runEl: Element, patch: RunPatch): void {
	if (runEl.localName === "br") return;
	const rPr = runProps(runEl);
	const doc = runEl.ownerDocument;

	if (patch.bold !== undefined) rPr.setAttribute("b", patch.bold ? "1" : "0");
	if (patch.italic !== undefined) rPr.setAttribute("i", patch.italic ? "1" : "0");
	if (patch.underline !== undefined) rPr.setAttribute("u", patch.underline ? "sng" : "none");
	if (patch.strike !== undefined) {
		rPr.setAttribute("strike", patch.strike ? "sngStrike" : "noStrike");
	}
	if (patch.size !== undefined) rPr.setAttribute("sz", String(Math.round(patch.size * 100)));

	if (patch.color !== undefined) {
		setOrderedChild(
			rPr,
			"solidFill",
			RPR_ORDER,
			patch.color === null ? null : solidFillElement(doc, patch.color),
		);
	}
	if (patch.font !== undefined) {
		let latin: Element | null = null;
		if (patch.font !== null) {
			latin = doc.createElementNS(A_NS, "a:latin");
			latin.setAttribute("typeface", patch.font);
		}
		setOrderedChild(rPr, "latin", RPR_ORDER, latin);
		// East Asian text falls back to a:ea, so keep the two in step or a
		// Japanese deck would keep rendering in the old face.
		let ea: Element | null = null;
		if (patch.font !== null) {
			ea = doc.createElementNS(A_NS, "a:ea");
			ea.setAttribute("typeface", patch.font);
		}
		setOrderedChild(rPr, "ea", RPR_ORDER, ea);
	}
}

/**
 * Split a run so that the characters in [start, end) live in a run of their own,
 * and return that run. Used when formatting applies to part of a run.
 */
export function splitRun(runEl: Element, start: number, end: number): Element | null {
	const t = child(runEl, "t");
	if (!t) return null;
	const text = t.textContent ?? "";
	const from = Math.max(0, Math.min(start, text.length));
	const to = Math.max(from, Math.min(end, text.length));
	if (to <= from) return null;
	if (from === 0 && to === text.length) return runEl;

	const parent = runEl.parentNode;
	if (!parent) return null;

	const middle = runEl.cloneNode(true) as Element;
	setRunText(middle, text.slice(from, to));

	if (to < text.length) {
		const tail = runEl.cloneNode(true) as Element;
		setRunText(tail, text.slice(to));
		parent.insertBefore(tail, runEl.nextSibling);
	}
	parent.insertBefore(middle, runEl.nextSibling);

	if (from > 0) setRunText(runEl, text.slice(0, from));
	else parent.removeChild(runEl);

	return middle;
}

function setRunText(runEl: Element, text: string): void {
	const t = child(runEl, "t");
	if (!t) return;
	t.textContent = text;
	if (/^\s|\s$/.test(text)) t.setAttribute("xml:space", "preserve");
	else t.removeAttribute("xml:space");
}

/** Apply a patch to every run in a text body. */
export function applyRunPatchToBody(txBody: Element, patch: RunPatch): boolean {
	let changed = false;
	for (const para of children(txBody, "p")) {
		for (const runEl of [...children(para, "r"), ...children(para, "fld")]) {
			applyRunPatch(runEl, patch);
			changed = true;
		}
		// An empty paragraph still has to remember the formatting, or typing into
		// it later would come back with the old size.
		const endPr = child(para, "endParaRPr");
		if (endPr) applyRunPatchOnProps(endPr, patch);
	}
	return changed;
}

function applyRunPatchOnProps(rPr: Element, patch: RunPatch): void {
	const doc = rPr.ownerDocument;
	if (patch.bold !== undefined) rPr.setAttribute("b", patch.bold ? "1" : "0");
	if (patch.italic !== undefined) rPr.setAttribute("i", patch.italic ? "1" : "0");
	if (patch.underline !== undefined) rPr.setAttribute("u", patch.underline ? "sng" : "none");
	if (patch.strike !== undefined) {
		rPr.setAttribute("strike", patch.strike ? "sngStrike" : "noStrike");
	}
	if (patch.size !== undefined) rPr.setAttribute("sz", String(Math.round(patch.size * 100)));
	if (patch.color !== undefined) {
		setOrderedChild(
			rPr,
			"solidFill",
			RPR_ORDER,
			patch.color === null ? null : solidFillElement(doc, patch.color),
		);
	}
}

// ------------------------------------------------------------ paragraphs

export interface ParagraphPatch {
	align?: "l" | "ctr" | "r" | "just";
	/** Outline level, 0-8. */
	level?: number;
	/** Relative level change, applied when `level` is absent. */
	levelDelta?: number;
	bullet?: "none" | "char" | "number";
	bulletChar?: string;
	/** Line spacing as a multiplier, e.g. 1.5. */
	lineSpacing?: number;
}

export function paragraphProps(para: Element): Element {
	const existing = child(para, "pPr");
	if (existing) return existing;
	const pPr = para.ownerDocument.createElementNS(A_NS, "a:pPr");
	para.insertBefore(pPr, para.firstChild);
	return pPr;
}

export function applyParagraphPatch(para: Element, patch: ParagraphPatch): void {
	const pPr = paragraphProps(para);
	const doc = para.ownerDocument;

	if (patch.align !== undefined) pPr.setAttribute("algn", patch.align);

	if (patch.level !== undefined || patch.levelDelta !== undefined) {
		const current = Number(pPr.getAttribute("lvl") ?? 0);
		const next = Math.max(
			0,
			Math.min(8, patch.level !== undefined ? patch.level : current + (patch.levelDelta ?? 0)),
		);
		if (next === 0) pPr.removeAttribute("lvl");
		else pPr.setAttribute("lvl", String(next));
		// Indentation comes from the level's list style, so any hand-set margin
		// would fight it; drop it and let the level decide.
		pPr.removeAttribute("marL");
		pPr.removeAttribute("indent");
	}

	if (patch.bullet !== undefined) {
		for (const name of ["buNone", "buAutoNum", "buChar"]) {
			setOrderedChild(pPr, name, PPR_ORDER, null);
		}
		if (patch.bullet === "none") {
			setOrderedChild(pPr, "buNone", PPR_ORDER, doc.createElementNS(A_NS, "a:buNone"));
		} else if (patch.bullet === "number") {
			const auto = doc.createElementNS(A_NS, "a:buAutoNum");
			auto.setAttribute("type", "arabicPeriod");
			setOrderedChild(pPr, "buAutoNum", PPR_ORDER, auto);
		} else {
			const bullet = doc.createElementNS(A_NS, "a:buChar");
			bullet.setAttribute("char", patch.bulletChar ?? "•");
			setOrderedChild(pPr, "buChar", PPR_ORDER, bullet);
		}
	}

	if (patch.lineSpacing !== undefined) {
		const lnSpc = doc.createElementNS(A_NS, "a:lnSpc");
		const pct = doc.createElementNS(A_NS, "a:spcPct");
		pct.setAttribute("val", String(Math.round(patch.lineSpacing * 100000)));
		lnSpc.appendChild(pct);
		setOrderedChild(pPr, "lnSpc", PPR_ORDER, lnSpc);
	}
}

export function applyParagraphPatchToBody(txBody: Element, patch: ParagraphPatch): boolean {
	const paragraphs = children(txBody, "p");
	for (const para of paragraphs) applyParagraphPatch(para, patch);
	return paragraphs.length > 0;
}

// ---------------------------------------------------------------- shapes

/** Set a shape's fill, or clear it with null. */
export function applyShapeFill(source: Element, color: string | null): boolean {
	const spPr = child(source, "spPr");
	if (!spPr) return false;
	const doc = source.ownerDocument;
	for (const name of ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"]) {
		setOrderedChild(spPr, name, SPPR_ORDER, null);
	}
	setOrderedChild(
		spPr,
		color === null ? "noFill" : "solidFill",
		SPPR_ORDER,
		color === null ? doc.createElementNS(A_NS, "a:noFill") : solidFillElement(doc, color),
	);
	return true;
}

export interface OutlinePatch {
	color?: string | null;
	/** Points. */
	width?: number;
	dash?: string | null;
}

export function applyShapeOutline(source: Element, patch: OutlinePatch): boolean {
	const spPr = child(source, "spPr");
	if (!spPr) return false;
	const doc = source.ownerDocument;

	if (patch.color === null) {
		const ln = doc.createElementNS(A_NS, "a:ln");
		ln.appendChild(doc.createElementNS(A_NS, "a:noFill"));
		setOrderedChild(spPr, "ln", SPPR_ORDER, ln);
		return true;
	}

	let ln = child(spPr, "ln");
	if (!ln) {
		ln = doc.createElementNS(A_NS, "a:ln");
		setOrderedChild(spPr, "ln", SPPR_ORDER, ln);
	}
	const LN_ORDER = [
		"noFill",
		"solidFill",
		"gradFill",
		"pattFill",
		"prstDash",
		"custDash",
		"round",
		"bevel",
		"miter",
		"headEnd",
		"tailEnd",
		"extLst",
	];

	if (patch.color !== undefined && patch.color !== null) {
		setOrderedChild(ln, "noFill", LN_ORDER, null);
		setOrderedChild(ln, "solidFill", LN_ORDER, solidFillElement(doc, patch.color));
	}
	if (patch.width !== undefined) {
		// a:ln@w is in EMU; a point is 12700 of them.
		ln.setAttribute("w", String(Math.round(patch.width * 12700)));
	}
	if (patch.dash !== undefined) {
		let dash: Element | null = null;
		if (patch.dash !== null) {
			dash = doc.createElementNS(A_NS, "a:prstDash");
			dash.setAttribute("val", patch.dash);
		}
		setOrderedChild(ln, "prstDash", LN_ORDER, dash);
	}
	return true;
}

/** Vertical anchor of a text body: t / ctr / b. */
export function applyTextAnchor(source: Element, anchor: "t" | "ctr" | "b"): boolean {
	const txBody = child(source, "txBody");
	const bodyPr = child(txBody, "bodyPr");
	if (!bodyPr) return false;
	bodyPr.setAttribute("anchor", anchor);
	return true;
}

/** Convert a run size in hundredths of a point to the pixels the model uses. */
export function sizeToPx(hundredths: number): number {
	return (hundredths / 100) * PT_TO_PX;
}
