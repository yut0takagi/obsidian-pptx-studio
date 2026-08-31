import type { Theme } from "./color";
import { DEFAULT_THEME, parseTheme } from "./color";
import { PptxPackage } from "./package";
import {
	findPlaceholder,
	type Inheritance,
	type InheritedShape,
	NO_INHERITANCE,
	parseBackground,
	parseShapeTree,
} from "./shapes";
import type { ParseContext } from "./style";
import { textBodyToPlain } from "./text";
import type { Deck, Fill, Shape, Slide } from "./types";
import { emuToPx } from "./types";
import { attr, boolAttr, child, children, numAttr } from "./xml";

const PRESENTATION = "ppt/presentation.xml";

/** Parse a .pptx into the renderer's model. Throws with a readable message on bad input. */
export function parseDeck(data: ArrayBuffer | Uint8Array, fallbackTitle: string): {
	deck: Deck;
	pkg: PptxPackage;
} {
	const pkg = PptxPackage.open(data);
	try {
		return { deck: rebuildDeck(pkg, fallbackTitle), pkg };
	} catch (e) {
		pkg.dispose();
		throw e;
	}
}

/**
 * Rebuild the model from an already-open package. Used after an edit: the XML
 * Documents are cached, so this only re-walks them — no unzip, no re-parse.
 */
export function rebuildDeck(pkg: PptxPackage, fallbackTitle: string): Deck {
	const presDoc = pkg.xml(PRESENTATION);
	const pres = presDoc?.documentElement ?? null;
	const sldSz = child(pres, "sldSz");
	const width = emuToPx(numAttr(sldSz, "cx") ?? 9144000);
	const height = emuToPx(numAttr(sldSz, "cy") ?? 6858000);
	const defaultTextStyle = child(pres, "defaultTextStyle");

	const themeCache = new Map<string, Theme>();
	const slides: Slide[] = [];

	const sldIds = children(child(pres, "sldIdLst"), "sldId");
	sldIds.forEach((sldId, i) => {
		// p:sldId carries both a numeric id and an r:id; only the latter is a relationship.
		const slidePath = pkg.relTarget(PRESENTATION, relIdOf(sldId));
		if (!slidePath || !pkg.has(slidePath)) return;
		try {
			slides.push(buildSlide(pkg, slidePath, i + 1, defaultTextStyle, themeCache));
		} catch (e) {
			slides.push(errorSlide(i + 1, slidePath, e as Error));
		}
	});

	return {
		width,
		height,
		slides,
		title: documentTitle(pkg) || fallbackTitle,
	};
}

function relIdOf(sldId: Element): string | null {
	for (const a of Array.from(sldId.attributes)) {
		if (a.localName === "id" && a.name !== "id") return a.value;
	}
	return null;
}

function documentTitle(pkg: PptxPackage): string {
	const doc = pkg.xml("docProps/core.xml");
	if (!doc?.documentElement) return "";
	for (let n = doc.documentElement.firstElementChild; n; n = n.nextElementSibling) {
		if (n.localName === "title") return (n.textContent ?? "").trim();
	}
	return "";
}

/**
 * Re-derive one slide in place.
 *
 * Most edits touch exactly one slide, and walking the other twenty-six to
 * rebuild them is work nobody asked for — on a big deck it is the difference
 * between an edit landing instantly and landing visibly late.
 */
export function rebuildSlideAt(pkg: PptxPackage, deck: Deck, index: number): boolean {
	const slide = deck.slides[index];
	if (!slide) return false;
	const pres = pkg.xml(PRESENTATION)?.documentElement ?? null;
	try {
		deck.slides[index] = buildSlide(
			pkg,
			slide.partPath,
			slide.index,
			child(pres, "defaultTextStyle"),
			new Map(),
		);
		return true;
	} catch {
		return false;
	}
}

function buildSlide(
	pkg: PptxPackage,
	slidePath: string,
	index: number,
	defaultTextStyle: Element | null,
	themeCache: Map<string, Theme>,
): Slide {
	const slideDoc = pkg.xml(slidePath);
	const slideRoot = slideDoc?.documentElement ?? null;
	const slideTree = child(child(slideRoot, "cSld"), "spTree");

	const layoutRel = pkg.relByKind(slidePath, "slideLayout");
	const layoutPath = layoutRel?.target ?? null;
	const layoutRoot = layoutPath ? (pkg.xml(layoutPath)?.documentElement ?? null) : null;
	const layoutTree = child(child(layoutRoot, "cSld"), "spTree");

	const masterRel = layoutPath ? pkg.relByKind(layoutPath, "slideMaster") : null;
	const masterPath = masterRel?.target ?? null;
	const masterRoot = masterPath ? (pkg.xml(masterPath)?.documentElement ?? null) : null;
	const masterTree = child(child(masterRoot, "cSld"), "spTree");

	const theme = resolveTheme(pkg, masterPath, masterRoot, themeCache);

	const slideCtx: ParseContext = { pkg, partPath: slidePath, theme };
	const layoutCtx: ParseContext = { pkg, partPath: layoutPath ?? slidePath, theme };
	const masterCtx: ParseContext = { pkg, partPath: masterPath ?? slidePath, theme };

	const txStyles = child(masterRoot, "txStyles");
	const inherit: Inheritance = {
		lookup(type, idx) {
			const out: InheritedShape[] = [];
			const layoutPh = findPlaceholder(layoutTree, type, idx);
			if (layoutPh) out.push({ sp: layoutPh, ctx: layoutCtx });
			const masterPh = findPlaceholder(masterTree, type, idx);
			if (masterPh) out.push({ sp: masterPh, ctx: masterCtx });
			return out;
		},
		masterStyles: {
			title: child(txStyles, "titleStyle"),
			body: child(txStyles, "bodyStyle"),
			other: child(txStyles, "otherStyle"),
		},
		defaultTextStyle,
	};

	// Master and layout decoration sits beneath the slide's own shapes. Their
	// placeholders are prompts, not content, so only non-placeholders are drawn.
	const showMasterOnLayout = boolAttr(layoutRoot, "showMasterSp") ?? true;
	const showLayoutOnSlide = boolAttr(slideRoot, "showMasterSp") ?? true;

	const shapes: Shape[] = [];
	if (showLayoutOnSlide && showMasterOnLayout) {
		shapes.push(...decorationShapes(masterTree, masterCtx));
	}
	if (showLayoutOnSlide) {
		shapes.push(...decorationShapes(layoutTree, layoutCtx));
	}
	const templateShapes = shapes.length;
	shapes.push(...parseShapeTree(slideTree, slideCtx, inherit));

	const background: Fill =
		parseBackground(child(child(slideRoot, "cSld"), "bg"), slideCtx) ??
		parseBackground(child(child(layoutRoot, "cSld"), "bg"), layoutCtx) ??
		parseBackground(child(child(masterRoot, "cSld"), "bg"), masterCtx) ??
		{ kind: "solid", color: "#ffffff" };

	return {
		index,
		name: attr(child(slideRoot, "cSld"), "name") ?? "",
		partPath: slidePath,
		background,
		shapes,
		templateShapes,
		notes: readNotes(pkg, slidePath, theme),
	};
}

/** Non-placeholder shapes from a layout or master, i.e. the template's own artwork. */
function decorationShapes(tree: Element | null, ctx: ParseContext): Shape[] {
	if (!tree) return [];
	const all = parseShapeTree(tree, ctx, NO_INHERITANCE);
	const placeholderIds = new Set<string>();
	for (const el of children(tree, "sp")) {
		const ph = child(child(child(el, "nvSpPr"), "nvPr"), "ph");
		if (ph) {
			const id = attr(child(child(el, "nvSpPr"), "cNvPr"), "id");
			if (id) placeholderIds.add(id);
		}
	}
	for (const el of children(tree, "pic")) {
		const ph = child(child(child(el, "nvPicPr"), "nvPr"), "ph");
		if (ph) {
			const id = attr(child(child(el, "nvPicPr"), "cNvPr"), "id");
			if (id) placeholderIds.add(id);
		}
	}
	return all.filter((s) => !placeholderIds.has(s.id));
}

function resolveTheme(
	pkg: PptxPackage,
	masterPath: string | null,
	masterRoot: Element | null,
	cache: Map<string, Theme>,
): Theme {
	if (!masterPath) return DEFAULT_THEME;
	const cached = cache.get(masterPath);
	if (cached) return cached;
	const themeRel = pkg.relByKind(masterPath, "theme");
	const themeRoot = themeRel ? (pkg.xml(themeRel.target)?.documentElement ?? null) : null;
	const theme = parseTheme(themeRoot, child(masterRoot, "clrMap"));
	cache.set(masterPath, theme);
	return theme;
}

function readNotes(pkg: PptxPackage, slidePath: string, theme: Theme): string {
	const rel = pkg.relByKind(slidePath, "notesSlide");
	if (!rel) return "";
	const root = pkg.xml(rel.target)?.documentElement ?? null;
	const tree = child(child(root, "cSld"), "spTree");
	if (!tree) return "";
	const ctx: ParseContext = { pkg, partPath: rel.target, theme };
	const shapes = parseShapeTree(tree, ctx, NO_INHERITANCE);
	const lines: string[] = [];
	for (const shape of shapes) {
		if (shape.kind !== "shape") continue;
		// Skip the slide-image and slide-number placeholders the notes master adds.
		if (shape.placeholder === "sldNum") continue;
		lines.push(...textBodyToPlain(shape.text));
	}
	return lines.join("\n").trim();
}

function errorSlide(index: number, path: string, error: Error): Slide {
	return {
		index,
		name: path,
		partPath: path,
		background: { kind: "solid", color: "#ffffff" },
		shapes: [],
		templateShapes: 0,
		notes: `Could not render this slide: ${error.message}`,
	};
}
