import type { ParseContext } from "./style";
import { highlightColor } from "./style";
import { colorContext } from "./style";
import { resolveFillColor } from "./color";
import type { Bullet, Paragraph, Run, TextBody } from "./types";
import { PT_TO_PX, emuToPx } from "./types";
import { attr, boolAttr, child, children, numAttr } from "./xml";

/**
 * The list-style elements that apply to a text body, highest priority first.
 * Each is an a:lstStyle (or p:titleStyle / p:bodyStyle / p:otherStyle) whose
 * a:lvlNpPr children supply defaults for paragraphs at level N.
 */
export type StyleChain = Element[];

const DEFAULT_INSETS: [number, number, number, number] = [
	emuToPx(91440),
	emuToPx(45720),
	emuToPx(91440),
	emuToPx(45720),
];

/** Body text default when nothing in the chain declares a size. */
const DEFAULT_SIZE_PT = 18;

export function parseTextBody(
	txBody: Element | null,
	ctx: ParseContext,
	chain: StyleChain,
	defaultColor: string | null,
): TextBody | null {
	if (!txBody) return null;

	const bodyPr = child(txBody, "bodyPr");
	const ownList = child(txBody, "lstStyle");
	const fullChain: StyleChain = ownList ? [ownList, ...chain] : chain;

	const normAutofit = child(bodyPr, "normAutofit");
	const anchorRaw = attr(bodyPr, "anchor");
	const vertRaw = attr(bodyPr, "vert");

	const paragraphs: Paragraph[] = [];
	// One counter per outline level, reset when a level restarts.
	const counters: number[] = [];
	for (const p of children(txBody, "p")) {
		paragraphs.push(parseParagraph(p, ctx, fullChain, defaultColor, counters));
	}

	return {
		source: txBody,
		sourcePart: ctx.partPath,
		anchor: anchorRaw === "ctr" ? "middle" : anchorRaw === "b" ? "bottom" : "top",
		insets: [
			readInset(bodyPr, "lIns", DEFAULT_INSETS[0]),
			readInset(bodyPr, "tIns", DEFAULT_INSETS[1]),
			readInset(bodyPr, "rIns", DEFAULT_INSETS[2]),
			readInset(bodyPr, "bIns", DEFAULT_INSETS[3]),
		],
		wrap: attr(bodyPr, "wrap") !== "none",
		fontScale: normAutofit ? (numAttr(normAutofit, "fontScale") ?? 100000) / 100000 : 1,
		lineSpaceReduction: normAutofit
			? (numAttr(normAutofit, "lnSpcReduction") ?? 0) / 100000
			: 0,
		paragraphs,
		vertical: vertRaw === "vert" ? "vert" : vertRaw === "vert270" ? "vert270" : "horz",
	};
}

function readInset(bodyPr: Element | null, name: string, fallback: number): number {
	const v = numAttr(bodyPr, name);
	return v === null ? fallback : emuToPx(v);
}

function parseParagraph(
	p: Element,
	ctx: ParseContext,
	chain: StyleChain,
	defaultColor: string | null,
	counters: number[],
): Paragraph {
	const pPr = child(p, "pPr");
	const level = Math.max(0, Math.min(8, numAttr(pPr, "lvl") ?? 0));
	// Level defaults from the style chain, highest priority first.
	const levelPrs = levelProperties(chain, level);
	const propSources = pPr ? [pPr, ...levelPrs] : levelPrs;

	const align = pickAttr(propSources, "algn");
	const marL = pickNum(propSources, "marL");
	const indent = pickNum(propSources, "indent");

	const runs: Run[] = [];
	const rPrSources = propSources.map((el) => child(el, "defRPr")).filter(isElement);
	for (let n = p.firstElementChild; n; n = n.nextElementSibling) {
		if (n.localName === "r") {
			const run = parseRun(n, ctx, rPrSources, defaultColor);
			if (run) runs.push(run);
		} else if (n.localName === "fld") {
			// Slide numbers, dates and similar: the cached <a:t> is what PowerPoint shows.
			const run = parseRun(n, ctx, rPrSources, defaultColor);
			if (run) runs.push(run);
		} else if (n.localName === "br") {
			const style = parseRunStyle(child(n, "rPr"), ctx, rPrSources, defaultColor);
			runs.push({ ...style, text: "\n", source: n });
		}
	}

	if (runs.length === 0) {
		// Keep empty paragraphs: they are how decks create vertical spacing. Size
		// them from a:endParaRPr so the gap matches PowerPoint's.
		const endStyle = parseRunStyle(child(p, "endParaRPr"), ctx, rPrSources, defaultColor);
		runs.push({ ...endStyle, text: "", source: null });
	}

	const bullet = parseBullet(propSources, ctx, level, runs[0], counters);

	return {
		source: p,
		level,
		align:
			align === "ctr"
				? "center"
				: align === "r"
					? "right"
					: align === "just" || align === "dist"
						? "justify"
						: "left",
		bullet,
		marginLeft: marL === null ? defaultMarginLeft(level, bullet) : emuToPx(marL),
		indent: indent === null ? defaultIndent(bullet) : emuToPx(indent),
		spaceBefore: spacingPx(propSources, "spcBef", runs[0]?.size ?? 0),
		spaceAfter: spacingPx(propSources, "spcAft", runs[0]?.size ?? 0),
		lineSpacing: lineSpacing(propSources),
		runs,
	};
}

function parseRun(
	r: Element,
	ctx: ParseContext,
	rPrSources: Element[],
	defaultColor: string | null,
): Run | null {
	const text = child(r, "t")?.textContent ?? "";
	const style = parseRunStyle(child(r, "rPr"), ctx, rPrSources, defaultColor);
	return { ...style, text, source: r };
}

function parseRunStyle(
	rPr: Element | null,
	ctx: ParseContext,
	inherited: Element[],
	defaultColor: string | null,
): Omit<Run, "text" | "source"> {
	const sources = rPr ? [rPr, ...inherited] : inherited;

	const szHundredths = pickNum(sources, "sz");
	const size = ((szHundredths ?? DEFAULT_SIZE_PT * 100) / 100) * PT_TO_PX;

	let color: string | null = null;
	for (const src of sources) {
		const solid = child(src, "solidFill");
		if (solid) {
			color = resolveFillColor(solid, colorContext(ctx));
			if (color) break;
		}
	}

	let font: string | null = null;
	for (const src of sources) {
		const typeface = attr(child(src, "latin"), "typeface");
		if (typeface) {
			font = resolveTypeface(typeface, ctx);
			break;
		}
	}

	let link: string | null = null;
	for (const src of sources) {
		const hlink = child(src, "hlinkClick");
		if (hlink) {
			const relId = attr(hlink, "id");
			const rel = relId ? ctx.pkg.rels(ctx.partPath).get(relId) : undefined;
			link = rel?.external ? rel.target : (rel?.target ?? null);
			break;
		}
	}

	return {
		size,
		bold: pickBool(sources, "b") ?? false,
		italic: pickBool(sources, "i") ?? false,
		underline: (pickAttr(sources, "u") ?? "none") !== "none",
		strike: (pickAttr(sources, "strike") ?? "noStrike") !== "noStrike",
		color: color ?? defaultColor ?? "#000000",
		font: font ?? ctx.theme.minorFont,
		baseline: (pickNum(sources, "baseline") ?? 0) / 1000,
		spacing: (pickNum(sources, "spc") ?? 0) / 100 * PT_TO_PX,
		link,
		highlight: rPr ? highlightColor(rPr, ctx) : null,
	};
}

/** "+mj-lt" / "+mn-ea" and friends point back at the theme's font scheme. */
function resolveTypeface(typeface: string, ctx: ParseContext): string {
	if (typeface.startsWith("+mj")) return ctx.theme.majorFont;
	if (typeface.startsWith("+mn")) return ctx.theme.minorFont;
	return typeface;
}

function parseBullet(
	sources: Element[],
	ctx: ParseContext,
	level: number,
	firstRun: Run | undefined,
	counters: number[],
): Bullet | null {
	let kind: "char" | "number" | "none" | null = null;
	let glyph = "";
	let autoNumType = "arabicPeriod";
	let startAt = 1;

	for (const src of sources) {
		if (child(src, "buNone")) {
			kind = "none";
			break;
		}
		const buChar = child(src, "buChar");
		if (buChar) {
			kind = "char";
			glyph = attr(buChar, "char") ?? "•";
			break;
		}
		const buAutoNum = child(src, "buAutoNum");
		if (buAutoNum) {
			kind = "number";
			autoNumType = attr(buAutoNum, "type") ?? "arabicPeriod";
			startAt = numAttr(buAutoNum, "startAt") ?? 1;
			break;
		}
	}

	if (kind === null || kind === "none") {
		// A level that stops numbering resets the counters below it.
		counters[level] = 0;
		return null;
	}

	let color: string | null = null;
	for (const src of sources) {
		const buClr = child(src, "buClr");
		if (buClr) {
			color = resolveFillColor(buClr, colorContext(ctx));
			break;
		}
	}

	let font: string | null = null;
	for (const src of sources) {
		const buFont = attr(child(src, "buFont"), "typeface");
		if (buFont) {
			font = buFont;
			break;
		}
	}

	let scale = 1;
	for (const src of sources) {
		const pct = numAttr(child(src, "buSzPct"), "val");
		if (pct !== null) {
			scale = pct / 100000;
			break;
		}
	}

	let text = glyph;
	if (kind === "number") {
		const current = (counters[level] ?? startAt - 1) + 1;
		counters[level] = current;
		for (let i = level + 1; i < counters.length; i++) counters[i] = 0;
		text = formatAutoNumber(current, autoNumType);
	} else {
		counters[level] = 0;
	}

	return { kind, text, color, font, scale };
}

function formatAutoNumber(n: number, type: string): string {
	const body = type.startsWith("alphaLc")
		? toAlpha(n).toLowerCase()
		: type.startsWith("alphaUc")
			? toAlpha(n)
			: type.startsWith("romanLc")
				? toRoman(n).toLowerCase()
				: type.startsWith("romanUc")
					? toRoman(n)
					: String(n);
	if (type.endsWith("ParenBoth")) return `(${body})`;
	if (type.endsWith("ParenR")) return `${body})`;
	if (type.endsWith("Period")) return `${body}.`;
	return body;
}

function toAlpha(n: number): string {
	let out = "";
	let v = n;
	while (v > 0) {
		const rem = (v - 1) % 26;
		out = String.fromCharCode(65 + rem) + out;
		v = Math.floor((v - 1) / 26);
	}
	return out || "A";
}

const ROMAN: [number, string][] = [
	[1000, "M"],
	[900, "CM"],
	[500, "D"],
	[400, "CD"],
	[100, "C"],
	[90, "XC"],
	[50, "L"],
	[40, "XL"],
	[10, "X"],
	[9, "IX"],
	[5, "V"],
	[4, "IV"],
	[1, "I"],
];

function toRoman(n: number): string {
	let v = n;
	let out = "";
	for (const [value, sym] of ROMAN) {
		while (v >= value) {
			out += sym;
			v -= value;
		}
	}
	return out || "I";
}

/** PowerPoint's built-in indents when a level declares none. */
function defaultMarginLeft(level: number, bullet: Bullet | null): number {
	if (!bullet) return level === 0 ? 0 : emuToPx(level * 457200);
	return emuToPx((level + 1) * 342900);
}

function defaultIndent(bullet: Bullet | null): number {
	return bullet ? emuToPx(-342900) : 0;
}

function spacingPx(sources: Element[], name: string, fontSize: number): number {
	for (const src of sources) {
		const spc = child(src, name);
		if (!spc) continue;
		const pct = numAttr(child(spc, "spcPct"), "val");
		if (pct !== null) return (pct / 100000) * fontSize;
		const pts = numAttr(child(spc, "spcPts"), "val");
		if (pts !== null) return (pts / 100) * PT_TO_PX;
	}
	return 0;
}

function lineSpacing(sources: Element[]): number | null {
	for (const src of sources) {
		const lnSpc = child(src, "lnSpc");
		if (!lnSpc) continue;
		const pct = numAttr(child(lnSpc, "spcPct"), "val");
		// PowerPoint's "single" spacing is tighter than a browser's default 1.2em.
		if (pct !== null) return (pct / 100000) * 1.2;
		const pts = numAttr(child(lnSpc, "spcPts"), "val");
		if (pts !== null) return -((pts / 100) * PT_TO_PX);
	}
	return null;
}

/** a:lvlNpPr for the given level, from each element of the chain. */
function levelProperties(chain: StyleChain, level: number): Element[] {
	const name = `lvl${level + 1}pPr`;
	const out: Element[] = [];
	for (const lst of chain) {
		const el = child(lst, name);
		if (el) out.push(el);
	}
	return out;
}

function pickAttr(sources: Element[], name: string): string | null {
	for (const src of sources) {
		const v = attr(src, name);
		if (v !== null) return v;
	}
	return null;
}

function pickNum(sources: Element[], name: string): number | null {
	for (const src of sources) {
		const v = numAttr(src, name);
		if (v !== null) return v;
	}
	return null;
}

function pickBool(sources: Element[], name: string): boolean | null {
	for (const src of sources) {
		const v = boolAttr(src, name);
		if (v !== null) return v;
	}
	return null;
}

function isElement(el: Element | null): el is Element {
	return el !== null;
}

/** Flatten a text body back to plain text, for Markdown extraction and search. */
export function textBodyToPlain(body: TextBody | null): string[] {
	if (!body) return [];
	return body.paragraphs.map((p) => p.runs.map((r) => r.text).join("").trim());
}
