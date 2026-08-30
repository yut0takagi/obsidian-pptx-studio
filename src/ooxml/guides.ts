import type { PptxPackage } from "../pptx/package";
import { attr, child, children } from "../pptx/xml";
import { CT, ensureOverride } from "./contentTypes";
import { REL_BASE, ensureRelationship, relsPathFor } from "./rels";
import { P_NS, p as pEl } from "./tree";

export const VIEW_PROPS = "ppt/viewProps.xml";
const PRESENTATION = "ppt/presentation.xml";

/**
 * Drawing guides, stored where PowerPoint stores them.
 *
 * `ppt/viewProps.xml` holds the guide list, so guides created here show up in
 * PowerPoint at the same positions and vice versa — they are a property of the
 * deck, not of this plugin.
 *
 * Positions are in eighths of a point, which works out at exactly six units per
 * CSS pixel at the 96dpi the renderer uses.
 */
const UNITS_PER_PX = 6;

export interface Guide {
	/** A "horz" guide is a horizontal line, so it fixes a y coordinate. */
	orientation: "horz" | "vert";
	/** Slide coordinate in pixels. */
	position: number;
}

const EMPTY_VIEW_PROPS =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
	`<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
	`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
	`xmlns:p="${P_NS}"><p:slideViewPr><p:cSldViewPr><p:cViewPr varScale="1">` +
	'<p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale>' +
	'<p:origin x="0" y="0"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr></p:viewPr>';

function guideList(pkg: PptxPackage, create: boolean): Element | null {
	let doc = pkg.xml(VIEW_PROPS);
	if (!doc?.documentElement) {
		if (!create) return null;
		pkg.replacePart(VIEW_PROPS, new TextEncoder().encode(EMPTY_VIEW_PROPS));
		ensureOverride(pkg, VIEW_PROPS, CT.viewProps);
		ensureRelationship(pkg, PRESENTATION, `${REL_BASE}/viewProps`, VIEW_PROPS);
		doc = pkg.xml(VIEW_PROPS);
	}
	const root = doc?.documentElement ?? null;
	const cSldViewPr = child(child(root, "slideViewPr"), "cSldViewPr");
	if (!cSldViewPr) {
		if (!create || !root) return null;
		// A viewProps that exists but has no slide view still needs somewhere to
		// put the list, and the schema wants it under p:cSldViewPr.
		const slideViewPr = pEl(root.ownerDocument, "slideViewPr");
		const created = pEl(root.ownerDocument, "cSldViewPr");
		const list = pEl(root.ownerDocument, "guideLst");
		created.appendChild(list);
		slideViewPr.appendChild(created);
		root.appendChild(slideViewPr);
		return list;
	}
	const existing = child(cSldViewPr, "guideLst");
	if (existing) return existing;
	if (!create) return null;
	const list = pEl(cSldViewPr.ownerDocument, "guideLst");
	cSldViewPr.appendChild(list);
	return list;
}

export function readGuides(pkg: PptxPackage): Guide[] {
	const list = guideList(pkg, false);
	if (!list) return [];
	return children(list, "guide").map((el) => ({
		// The attribute is optional and defaults to a horizontal guide.
		orientation: attr(el, "orient") === "vert" ? "vert" : "horz",
		position: Number(attr(el, "pos") ?? 0) / UNITS_PER_PX,
	}));
}

/** Replace the whole guide list. Returns false when nothing changed. */
export function writeGuides(pkg: PptxPackage, guides: Guide[]): boolean {
	const list = guideList(pkg, true);
	if (!list) return false;

	const before = readGuides(pkg);
	if (
		before.length === guides.length &&
		before.every(
			(g, i) =>
				g.orientation === guides[i].orientation &&
				Math.abs(g.position - guides[i].position) < 0.01,
		)
	) {
		return false;
	}

	while (list.firstChild) list.removeChild(list.firstChild);
	for (const guide of guides) {
		const el = pEl(list.ownerDocument, "guide");
		if (guide.orientation === "vert") el.setAttribute("orient", "vert");
		el.setAttribute("pos", String(Math.round(guide.position * UNITS_PER_PX)));
		list.appendChild(el);
	}
	pkg.markDirty(VIEW_PROPS);
	return true;
}

/** Parts a guide edit touches, for the undo snapshot. */
export function guideParts(): string[] {
	return [VIEW_PROPS, relsPathFor(PRESENTATION), "[Content_Types].xml"];
}
