import { emptyChart, parseChartPart } from "./chart";
import { resolveFillColor } from "./color";
import { EMPTY_FRAME, geometryName, isLineGeometry, parseChildFrame, parseFrame } from "./geometry";
import type { ParseContext } from "./style";
import {
	colorContext,
	colorFromFontRef,
	fillFromStyleRef,
	parseCrop,
	parseFill,
	parseStroke,
	strokeFromStyleRef,
} from "./style";
import type { StyleChain } from "./text";
import { parseTextBody } from "./text";
import type { Fill, Frame, Shape, Table, TableCell, TextBody } from "./types";
import { emuToPx } from "./types";
import { attr, boolAttr, child, children, numAttr } from "./xml";

/** A placeholder inherited from a layout or master, with the part it came from. */
export interface InheritedShape {
	sp: Element;
	ctx: ParseContext;
}

export interface Inheritance {
	/** Layout then master placeholders that match a slide placeholder, in that order. */
	lookup(type: string | null, idx: string | null): InheritedShape[];
	/** p:txStyles from the master, keyed by placeholder family. */
	masterStyles: { title: Element | null; body: Element | null; other: Element | null };
	/** p:defaultTextStyle from presentation.xml. */
	defaultTextStyle: Element | null;
}

export const NO_INHERITANCE: Inheritance = {
	lookup: () => [],
	masterStyles: { title: null, body: null, other: null },
	defaultTextStyle: null,
};

/** Parse a p:spTree (or dsp:spTree) into renderable shapes. */
export function parseShapeTree(
	spTree: Element | null,
	ctx: ParseContext,
	inherit: Inheritance,
): Shape[] {
	if (!spTree) return [];
	const shapes: Shape[] = [];
	for (const el of renderableChildren(spTree)) {
		const shape = parseShape(el, ctx, inherit);
		if (shape) shapes.push(shape);
	}
	return shapes;
}

/**
 * Children of a shape tree, unwrapping mc:AlternateContent. The Fallback branch
 * is preferred: it is the one written for consumers that do not understand the
 * newer markup in mc:Choice, which is exactly our situation.
 */
function renderableChildren(parent: Element): Element[] {
	const out: Element[] = [];
	for (let n = parent.firstElementChild; n; n = n.nextElementSibling) {
		if (n.localName === "AlternateContent") {
			const fallback = child(n, "Fallback");
			const choice = child(n, "Choice");
			const branch = fallback && fallback.firstElementChild ? fallback : choice;
			if (branch) out.push(...renderableChildren(branch));
			continue;
		}
		out.push(n);
	}
	return out;
}

function parseShape(el: Element, ctx: ParseContext, inherit: Inheritance): Shape | null {
	switch (el.localName) {
		case "sp":
			return parseSp(el, ctx, inherit);
		case "pic":
			return parsePic(el, ctx, inherit);
		case "graphicFrame":
			return parseGraphicFrame(el, ctx);
		case "grpSp":
			return parseGroup(el, ctx, inherit);
		case "cxnSp":
			return parseConnector(el, ctx);
		default:
			return null;
	}
}

interface Identity {
	id: string;
	name: string;
	hidden: boolean;
}

function identity(el: Element, nvContainer: string): Identity {
	const cNvPr = child(child(el, nvContainer), "cNvPr");
	return {
		id: attr(cNvPr, "id") ?? "",
		name: attr(cNvPr, "name") ?? "",
		hidden: boolAttr(cNvPr, "hidden") ?? false,
	};
}

function placeholderRef(el: Element, nvContainer: string): { type: string | null; idx: string | null } | null {
	const ph = child(child(child(el, nvContainer), "nvPr"), "ph");
	if (!ph) return null;
	return { type: attr(ph, "type"), idx: attr(ph, "idx") };
}

// ------------------------------------------------------------------ p:sp

function parseSp(el: Element, ctx: ParseContext, inherit: Inheritance): Shape | null {
	const ident = identity(el, "nvSpPr");
	const ph = placeholderRef(el, "nvSpPr");
	const inherited = ph ? inherit.lookup(ph.type, ph.idx) : [];

	const spPr = child(el, "spPr");
	const style = child(el, "style");

	const frame =
		parseFrame(child(spPr, "xfrm")) ??
		firstNonNull(inherited, (src) => parseFrame(childPathXfrm(src.sp))) ??
		EMPTY_FRAME;

	const geom = geometryName(spPr) ||
		firstNonNull(inherited, (src) => geometryName(child(src.sp, "spPr"))) ||
		"rect";

	const fill =
		parseFill(spPr, ctx) ??
		fillFromStyleRef(style, ctx) ??
		firstNonNull(inherited, (src) => parseFill(child(src.sp, "spPr"), src.ctx)) ??
		firstNonNull(inherited, (src) => fillFromStyleRef(child(src.sp, "style"), src.ctx));

	const stroke =
		parseStroke(spPr, ctx) ??
		strokeFromStyleRef(style, ctx) ??
		firstNonNull(inherited, (src) => parseStroke(child(src.sp, "spPr"), src.ctx)) ??
		firstNonNull(inherited, (src) => strokeFromStyleRef(child(src.sp, "style"), src.ctx));

	const defaultColor =
		colorFromFontRef(style, ctx) ??
		firstNonNull(inherited, (src) => colorFromFontRef(child(src.sp, "style"), src.ctx));

	const chain = buildStyleChain(ph?.type ?? null, inherited, inherit);
	const text = parseTextBody(child(el, "txBody"), ctx, chain, defaultColor);

	if (isLineGeometry(geom) && !hasVisibleText(text)) {
		return {
			kind: "line",
			id: ident.id,
			name: ident.name,
			hidden: ident.hidden,
			source: el,
			sourcePart: ctx.partPath,
			frame,
			stroke,
			headArrow: false,
			tailArrow: false,
		};
	}

	return {
		kind: "shape",
		id: ident.id,
		name: ident.name,
		hidden: ident.hidden,
		source: el,
		sourcePart: ctx.partPath,
		frame,
		geom,
		fill,
		stroke,
		text,
		placeholder: ph?.type ?? (ph ? "body" : null),
	};
}

function childPathXfrm(sp: Element): Element | null {
	return child(child(sp, "spPr"), "xfrm");
}

function hasVisibleText(text: TextBody | null): boolean {
	if (!text) return false;
	return text.paragraphs.some((p) => p.runs.some((r) => r.text.trim() !== ""));
}

/**
 * The text style chain for a placeholder: its layout placeholder, then its
 * master placeholder, then the master's family-wide styles, then the
 * presentation default.
 */
function buildStyleChain(
	phType: string | null,
	inherited: InheritedShape[],
	inherit: Inheritance,
): StyleChain {
	const chain: StyleChain = [];
	for (const src of inherited) {
		const lst = child(child(src.sp, "txBody"), "lstStyle");
		if (lst) chain.push(lst);
	}
	const family = placeholderFamily(phType);
	const masterStyle =
		family === "title"
			? inherit.masterStyles.title
			: family === "body"
				? inherit.masterStyles.body
				: inherit.masterStyles.other;
	if (masterStyle) chain.push(masterStyle);
	if (inherit.defaultTextStyle) chain.push(inherit.defaultTextStyle);
	return chain;
}

/** Which of the master's three style blocks governs a placeholder type. */
export function placeholderFamily(type: string | null): "title" | "body" | "other" {
	switch (type) {
		case "title":
		case "ctrTitle":
			return "title";
		case "body":
		case "subTitle":
		case "obj":
		case null:
		case undefined:
			return "body";
		default:
			return "other";
	}
}

// ----------------------------------------------------------------- p:pic

function parsePic(el: Element, ctx: ParseContext, inherit: Inheritance): Shape | null {
	const ident = identity(el, "nvPicPr");
	const ph = placeholderRef(el, "nvPicPr");
	const inherited = ph ? inherit.lookup(ph.type, ph.idx) : [];

	const spPr = child(el, "spPr");
	const blipFill = child(el, "blipFill");
	const blip = child(blipFill, "blip");
	const mediaPath = ctx.pkg.relTarget(ctx.partPath, attr(blip, "embed"));

	const frame =
		parseFrame(child(spPr, "xfrm")) ??
		firstNonNull(inherited, (src) => parseFrame(childPathXfrm(src.sp))) ??
		EMPTY_FRAME;

	const cNvPr = child(child(el, "nvPicPr"), "cNvPr");
	const label = attr(cNvPr, "descr") || ident.name || "Image";

	return {
		kind: "image",
		id: ident.id,
		name: ident.name,
		hidden: ident.hidden,
		source: el,
		sourcePart: ctx.partPath,
		frame,
		url: ctx.pkg.mediaUrl(mediaPath),
		mediaPath,
		crop: parseCrop(blipFill),
		label,
		geom: geometryName(spPr),
		stroke: parseStroke(spPr, ctx),
	};
}

// -------------------------------------------------------------- p:grpSp

function parseGroup(el: Element, ctx: ParseContext, inherit: Inheritance): Shape | null {
	const ident = identity(el, "nvGrpSpPr");
	const xfrm = child(child(el, "grpSpPr"), "xfrm");
	const frame = parseFrame(xfrm) ?? EMPTY_FRAME;
	const childOffset = parseChildFrame(xfrm) ?? { x: 0, y: 0, w: frame.w, h: frame.h };

	const kids: Shape[] = [];
	for (const kid of renderableChildren(el)) {
		const shape = parseShape(kid, ctx, inherit);
		if (shape) kids.push(shape);
	}
	if (kids.length === 0) return null;

	return {
		kind: "group",
		id: ident.id,
		name: ident.name,
		hidden: ident.hidden,
		source: el,
		sourcePart: ctx.partPath,
		frame,
		childOffset,
		children: kids,
	};
}

// -------------------------------------------------------------- p:cxnSp

function parseConnector(el: Element, ctx: ParseContext): Shape | null {
	const ident = identity(el, "nvCxnSpPr");
	const spPr = child(el, "spPr");
	const frame = parseFrame(child(spPr, "xfrm")) ?? EMPTY_FRAME;
	const ln = child(spPr, "ln");
	return {
		kind: "line",
		id: ident.id,
		name: ident.name,
		hidden: ident.hidden,
		source: el,
		sourcePart: ctx.partPath,
		frame,
		stroke: parseStroke(spPr, ctx) ?? strokeFromStyleRef(child(el, "style"), ctx),
		headArrow: child(ln, "headEnd") !== null && attr(child(ln, "headEnd"), "type") !== "none",
		tailArrow: child(ln, "tailEnd") !== null && attr(child(ln, "tailEnd"), "type") !== "none",
	};
}

// ------------------------------------------------------ p:graphicFrame

const URI_TABLE = "table";
const URI_CHART = "chart";
const URI_DIAGRAM = "diagram";

function parseGraphicFrame(el: Element, ctx: ParseContext): Shape | null {
	const ident = identity(el, "nvGraphicFramePr");
	const frame = parseFrame(child(el, "xfrm")) ?? EMPTY_FRAME;
	const data = child(child(el, "graphic"), "graphicData");
	if (!data) return null;
	const uri = attr(data, "uri") ?? "";
	const base = {
		id: ident.id,
		name: ident.name,
		hidden: ident.hidden,
		source: el,
		sourcePart: ctx.partPath,
		frame,
	};

	if (uri.endsWith(URI_TABLE)) {
		const tbl = child(data, "tbl");
		if (!tbl) return null;
		return { kind: "table", ...base, table: parseTable(tbl, ctx, frame) };
	}

	if (uri.endsWith(URI_CHART)) {
		const path = ctx.pkg.relTarget(ctx.partPath, attr(child(data, "chart"), "id"));
		return {
			kind: "chart",
			...base,
			chart: path ? parseChartPart(ctx.pkg, path, ctx) : emptyChart(),
		};
	}

	if (uri.endsWith(URI_DIAGRAM)) {
		// SmartArt ships a plain-shapes rendering alongside the semantic model;
		// that fallback is what PowerPoint itself draws, so we reuse it.
		const group = parseDiagram(data, ctx, frame, ident);
		if (group) return group;
		return null;
	}

	return null;
}

function parseTable(tbl: Element, ctx: ParseContext, frame: Frame): Table {
	const grid = children(child(tbl, "tblGrid"), "gridCol").map((c) => emuToPx(numAttr(c, "w") ?? 0));
	const totalWidth = grid.reduce((a, b) => a + b, 0);
	// Column widths occasionally disagree with the frame; scale them to fit.
	const columns =
		totalWidth > 0 && frame.w > 0 ? grid.map((w) => (w / totalWidth) * frame.w) : grid;

	const rows: Table["rows"] = [];
	const covered = new Set<string>();

	const trs = children(tbl, "tr");
	trs.forEach((tr, rowIndex) => {
		const cells: TableCell[] = [];
		children(tr, "tc").forEach((tc, colIndex) => {
			const key = `${rowIndex},${colIndex}`;
			const hMerge = boolAttr(tc, "hMerge") ?? false;
			const vMerge = boolAttr(tc, "vMerge") ?? false;
			const colSpan = numAttr(tc, "gridSpan") ?? 1;
			const rowSpan = numAttr(tc, "rowSpan") ?? 1;
			const merged = hMerge || vMerge || covered.has(key);

			for (let r = 0; r < rowSpan; r++) {
				for (let c = 0; c < colSpan; c++) {
					if (r || c) covered.add(`${rowIndex + r},${colIndex + c}`);
				}
			}

			const tcPr = child(tc, "tcPr");
			const anchorRaw = attr(tcPr, "anchor");
			cells.push({
				text: parseTextBody(child(tc, "txBody"), ctx, [], null),
				fill: parseFill(tcPr, ctx),
				borders: {
					left: parseStroke(wrapLine(tcPr, "lnL"), ctx),
					top: parseStroke(wrapLine(tcPr, "lnT"), ctx),
					right: parseStroke(wrapLine(tcPr, "lnR"), ctx),
					bottom: parseStroke(wrapLine(tcPr, "lnB"), ctx),
				},
				colSpan,
				rowSpan,
				merged,
				anchor: anchorRaw === "ctr" ? "middle" : anchorRaw === "b" ? "bottom" : "top",
				margins: [
					emuToPx(numAttr(tcPr, "marL") ?? 91440),
					emuToPx(numAttr(tcPr, "marT") ?? 45720),
					emuToPx(numAttr(tcPr, "marR") ?? 91440),
					emuToPx(numAttr(tcPr, "marB") ?? 45720),
				],
			});
		});
		rows.push({ height: emuToPx(numAttr(tr, "h") ?? 0), cells });
	});

	return { columns, rows };
}

/**
 * Cell borders live in a:lnL/lnT/lnR/lnB rather than a:ln, so present them to
 * parseStroke under the name it expects.
 */
function wrapLine(tcPr: Element | null, name: string): Element | null {
	const ln = child(tcPr, name);
	if (!ln) return null;
	const holder = ln.ownerDocument.createElementNS(ln.namespaceURI, "a:__lnHolder");
	const clone = ln.cloneNode(true) as Element;
	const renamed = ln.ownerDocument.createElementNS(ln.namespaceURI, "a:ln");
	for (const a of Array.from(clone.attributes)) renamed.setAttribute(a.name, a.value);
	while (clone.firstChild) renamed.appendChild(clone.firstChild);
	holder.appendChild(renamed);
	return holder;
}

function parseDiagram(
	data: Element,
	ctx: ParseContext,
	frame: Frame,
	ident: Identity,
): Shape | null {
	const relIds = child(data, "relIds");
	const dataPath = ctx.pkg.relTarget(ctx.partPath, attr(relIds, "dm"));
	if (!dataPath) return null;
	const drawingRel = ctx.pkg.relByKind(dataPath, "diagramDrawing");
	if (!drawingRel) return null;
	const doc = ctx.pkg.xml(drawingRel.target);
	if (!doc?.documentElement) return null;

	const spTree = child(doc.documentElement, "spTree");
	if (!spTree) return null;

	const drawingCtx: ParseContext = { ...ctx, partPath: drawingRel.target };
	const kids = parseShapeTree(spTree, drawingCtx, NO_INHERITANCE);
	if (kids.length === 0) return null;

	// The drawing uses the frame's own coordinate space, so children map 1:1.
	const groupXfrm = child(spTree, "grpSpPr");
	const childFrame = parseChildFrame(child(groupXfrm, "xfrm")) ?? {
		x: 0,
		y: 0,
		w: frame.w,
		h: frame.h,
	};

	return {
		kind: "group",
		id: ident.id,
		name: ident.name || "SmartArt",
		hidden: ident.hidden,
		source: null,
		sourcePart: drawingRel.target,
		frame,
		childOffset: childFrame.w > 0 && childFrame.h > 0 ? childFrame : { x: 0, y: 0, w: frame.w, h: frame.h },
		children: kids,
	};
}

// -------------------------------------------------------------- helpers

function firstNonNull<T, R>(items: T[], fn: (item: T) => R | null): R | null {
	for (const item of items) {
		const value = fn(item);
		if (value !== null && value !== undefined) return value;
	}
	return null;
}

/** Slide/layout/master background, which lives in p:bg rather than on a shape. */
export function parseBackground(bg: Element | null, ctx: ParseContext): Fill | null {
	if (!bg) return null;
	const bgPr = child(bg, "bgPr");
	if (bgPr) return parseFill(bgPr, ctx);
	const bgRef = child(bg, "bgRef");
	if (bgRef) {
		const idx = numAttr(bgRef, "idx") ?? 0;
		const phClr = resolveFillColor(bgRef, colorContext(ctx)) ?? undefined;
		const source =
			idx >= 1001 ? ctx.theme.fmt.bgFillStyles[idx - 1001] : ctx.theme.fmt.fillStyles[idx - 1];
		if (!source) return phClr ? { kind: "solid", color: phClr } : null;
		const holder = source.ownerDocument.createElement("holder");
		holder.appendChild(source.cloneNode(true));
		return parseFill(holder, ctx, phClr);
	}
	return null;
}

/** Find the placeholder in a layout/master shape tree that a slide placeholder inherits from. */
export function findPlaceholder(
	spTree: Element | null,
	type: string | null,
	idx: string | null,
): Element | null {
	if (!spTree) return null;
	const candidates: { el: Element; type: string | null; idx: string | null }[] = [];
	const collect = (parent: Element) => {
		for (const el of renderableChildren(parent)) {
			if (el.localName === "sp" || el.localName === "pic" || el.localName === "graphicFrame") {
				const container =
					el.localName === "sp"
						? "nvSpPr"
						: el.localName === "pic"
							? "nvPicPr"
							: "nvGraphicFramePr";
				const ph = placeholderRef(el, container);
				if (ph) candidates.push({ el, type: ph.type, idx: ph.idx });
			}
			if (el.localName === "grpSp") collect(el);
		}
	};
	collect(spTree);

	// An explicit idx match wins; PowerPoint only falls back to type matching.
	if (idx !== null) {
		const byIdx = candidates.find((c) => c.idx === idx);
		if (byIdx) return byIdx.el;
	}
	const family = placeholderFamily(type);
	const byType = candidates.find((c) => placeholderFamily(c.type) === family && c.type === type);
	if (byType) return byType.el;
	const byFamily = candidates.find((c) => placeholderFamily(c.type) === family);
	return byFamily?.el ?? null;
}
