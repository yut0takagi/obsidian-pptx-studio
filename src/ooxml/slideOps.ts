import type { PptxPackage } from "../pptx/package";
import { attr, child, children } from "../pptx/xml";
import { CONTENT_TYPES_PART, CT, ensureOverride, removeOverride } from "./contentTypes";
import { REL_BASE, ensureRelationship, relsPathFor, removeRelationship } from "./rels";
import { P_NS, nonVisualProps, p as pEl, shapeElements, spTreeOf } from "./tree";

const PRESENTATION = "ppt/presentation.xml";
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/** Every part a slide-level operation may touch. */
export function slideOpParts(pkg: PptxPackage, extra: string[] = []): string[] {
	return [PRESENTATION, relsPathFor(PRESENTATION), CONTENT_TYPES_PART, ...extra];
}

function presentationRoot(pkg: PptxPackage): Element | null {
	return pkg.xml(PRESENTATION)?.documentElement ?? null;
}

function slideIdList(pkg: PptxPackage): Element | null {
	return child(presentationRoot(pkg), "sldIdLst");
}

/** Slide part paths in presentation order. */
export function slidePaths(pkg: PptxPackage): string[] {
	const list = slideIdList(pkg);
	const out: string[] = [];
	for (const sldId of children(list, "sldId")) {
		const relId = relationshipId(sldId);
		const path = relId ? pkg.relTarget(PRESENTATION, relId) : null;
		if (path) out.push(path);
	}
	return out;
}

function relationshipId(sldId: Element): string | null {
	for (const a of Array.from(sldId.attributes)) {
		if (a.localName === "id" && a.name !== "id") return a.value;
	}
	return null;
}

/** The part path the next created slide will take, known before it is created. */
export function nextSlidePath(pkg: PptxPackage): string {
	return `ppt/slides/slide${nextSlideNumber(pkg)}.xml`;
}

function nextSlideNumber(pkg: PptxPackage): number {
	let max = 0;
	for (const path of pkg.partPaths()) {
		const match = /^ppt\/slides\/slide(\d+)\.xml$/.exec(path);
		if (match) max = Math.max(max, Number(match[1]));
	}
	return max + 1;
}

function nextSlideId(list: Element): number {
	let max = 255;
	for (const sldId of children(list, "sldId")) {
		const value = Number(attr(sldId, "id") ?? 0);
		if (Number.isFinite(value)) max = Math.max(max, value);
	}
	return max + 1;
}

/** Insert a slide reference into the presentation at a given position. */
function registerSlide(pkg: PptxPackage, slidePath: string, position: number): void {
	const list = slideIdList(pkg);
	if (!list) throw new Error("This presentation has no slide list.");
	const relId = ensureRelationship(pkg, PRESENTATION, `${REL_BASE}/slide`, slidePath);
	const sldId = pEl(list.ownerDocument, "sldId");
	sldId.setAttribute("id", String(nextSlideId(list)));
	sldId.setAttributeNS(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		"r:id",
		relId,
	);
	const existing = children(list, "sldId");
	const anchor = existing[position] ?? null;
	list.insertBefore(sldId, anchor);
	ensureOverride(pkg, slidePath, CT.slide);
	pkg.markDirty(PRESENTATION);
}

/**
 * Create a slide from a layout.
 *
 * The layout's placeholders are copied across empty, which is what makes the new
 * slide show a "click to add title" box rather than a blank rectangle: a slide
 * placeholder inherits its position and styling from the matching layout one.
 */
export function addSlide(
	pkg: PptxPackage,
	layoutPath: string,
	position: number,
	slidePath = nextSlidePath(pkg),
): string {

	const layoutRoot = pkg.xml(layoutPath)?.documentElement ?? null;
	const layoutTree = child(child(layoutRoot, "cSld"), "spTree");

	const placeholders: string[] = [];
	let id = 2;
	for (const el of shapeElements(layoutTree)) {
		const ph = child(child(child(el, "nvSpPr"), "nvPr"), "ph");
		if (!ph) continue;
		const type = attr(ph, "type") ?? "body";
		// Dates, footers and slide numbers come from the layout itself; copying
		// them onto the slide would double them up.
		if (type === "dt" || type === "ftr" || type === "sldNum") continue;
		const idx = attr(ph, "idx");
		const name = nonVisualProps(el)?.getAttribute("name") ?? type;
		placeholders.push(
			`<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="${escapeXml(name)}"/><p:cNvSpPr>` +
				'<a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="' +
				`${escapeXml(type)}"${idx ? ` idx="${escapeXml(idx)}"` : ""}/></p:nvPr></p:nvSpPr>` +
				"<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>",
		);
	}

	const xml =
		DECL +
		'<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
		'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
		'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
		'<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>' +
		'</p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
		'<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
		placeholders.join("") +
		"</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>";

	pkg.addPart(slidePath, new TextEncoder().encode(xml));
	ensureRelationship(pkg, slidePath, `${REL_BASE}/slideLayout`, layoutPath);
	registerSlide(pkg, slidePath, position);
	return slidePath;
}

/** Copy a slide, including its relationships, and insert it after the original. */
export function duplicateSlide(
	pkg: PptxPackage,
	sourcePath: string,
	position: number,
	slidePath = nextSlidePath(pkg),
): string {
	const bytes = pkg.serializePart(sourcePath);
	if (!bytes) throw new Error(`Slide not found: ${sourcePath}`);
	pkg.addPart(slidePath, bytes);

	// Relationship targets are relative to the slides folder, which the copy
	// shares, so the original .rels can be reused verbatim.
	const sourceRels = pkg.serializePart(relsPathFor(sourcePath));
	if (sourceRels) pkg.replacePart(relsPathFor(slidePath), sourceRels);

	// A notes slide belongs to exactly one slide, so drop the copy's link to it
	// rather than leaving two slides claiming the same notes.
	for (const [id, rel] of pkg.rels(slidePath)) {
		if (rel.kind === "notesSlide") removeRelationship(pkg, slidePath, id);
	}

	registerSlide(pkg, slidePath, position);
	return slidePath;
}

/** Remove a slide and everything that existed only to serve it. */
export function deleteSlide(pkg: PptxPackage, slidePath: string): boolean {
	const list = slideIdList(pkg);
	if (!list) return false;
	const entries = children(list, "sldId");
	if (entries.length <= 1) return false;

	let removed = false;
	for (const sldId of entries) {
		const relId = relationshipId(sldId);
		if (!relId) continue;
		if (pkg.relTarget(PRESENTATION, relId) !== slidePath) continue;
		list.removeChild(sldId);
		removeRelationship(pkg, PRESENTATION, relId);
		removed = true;
		break;
	}
	if (!removed) return false;

	const notes = pkg.relByKind(slidePath, "notesSlide");
	if (notes) {
		removeOverride(pkg, notes.target);
		pkg.replacePart(relsPathFor(notes.target), null);
		pkg.replacePart(notes.target, null);
	}

	pkg.replacePart(relsPathFor(slidePath), null);
	pkg.replacePart(slidePath, null);
	removeOverride(pkg, slidePath);
	pkg.markDirty(PRESENTATION);
	return true;
}

/** Move a slide to a new position in the deck. */
export function moveSlide(pkg: PptxPackage, from: number, to: number): boolean {
	const list = slideIdList(pkg);
	if (!list) return false;
	const entries = children(list, "sldId");
	if (from < 0 || from >= entries.length) return false;
	const clamped = Math.max(0, Math.min(entries.length - 1, to));
	if (clamped === from) return false;

	const moving = entries[from];
	list.removeChild(moving);
	const rest = children(list, "sldId");
	list.insertBefore(moving, rest[clamped] ?? null);
	pkg.markDirty(PRESENTATION);
	return true;
}

/**
 * Set or clear a slide's own background.
 *
 * Clearing removes p:bg entirely rather than writing white, so the slide goes
 * back to inheriting whatever its layout and master specify.
 */
export function setSlideBackground(
	pkg: PptxPackage,
	slidePath: string,
	color: string | null,
): boolean {
	const root = pkg.xml(slidePath)?.documentElement ?? null;
	const cSld = child(root, "cSld");
	if (!cSld) return false;
	const doc = cSld.ownerDocument;

	const existing = child(cSld, "bg");
	if (existing) cSld.removeChild(existing);
	if (color === null) {
		pkg.markDirty(slidePath);
		return true;
	}

	const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
	const bg = pEl(doc, "bg");
	const bgPr = pEl(doc, "bgPr");
	const fill = doc.createElementNS(A, "a:solidFill");
	const clr = doc.createElementNS(A, "a:srgbClr");
	clr.setAttribute("val", color.replace(/^#/, "").toUpperCase());
	fill.appendChild(clr);
	bgPr.appendChild(fill);
	bgPr.appendChild(doc.createElementNS(A, "a:effectLst"));
	bg.appendChild(bgPr);
	// p:bg must be the first child of p:cSld.
	cSld.insertBefore(bg, cSld.firstChild);
	pkg.markDirty(slidePath);
	return true;
}

/** Layouts available for a new slide, in the order the master lists them. */
export function availableLayouts(pkg: PptxPackage): { path: string; name: string }[] {
	const root = presentationRoot(pkg);
	const masterId = children(child(root, "sldMasterIdLst"), "sldMasterId")[0];
	const masterRel = masterId ? relationshipId(masterId) : null;
	const masterPath = masterRel ? pkg.relTarget(PRESENTATION, masterRel) : null;
	if (!masterPath) return [];

	const masterRoot = pkg.xml(masterPath)?.documentElement ?? null;
	const out: { path: string; name: string }[] = [];
	for (const layoutId of children(child(masterRoot, "sldLayoutIdLst"), "sldLayoutId")) {
		const relId = relationshipId(layoutId);
		const path = relId ? pkg.relTarget(masterPath, relId) : null;
		if (!path) continue;
		const layoutRoot = pkg.xml(path)?.documentElement ?? null;
		const name = attr(child(layoutRoot, "cSld"), "name") ?? path.split("/").pop() ?? path;
		out.push({ path, name });
	}
	return out;
}

/** The layout a slide is built on, so a new slide can default to the same one. */
export function layoutOf(pkg: PptxPackage, slidePath: string): string | null {
	return pkg.relByKind(slidePath, "slideLayout")?.target ?? null;
}

/** True when the slide has any content of its own, used to warn before deleting. */
export function slideHasContent(pkg: PptxPackage, slidePath: string): boolean {
	return shapeElements(spTreeOf(pkg, slidePath)).length > 0;
}

function escapeXml(value: string): string {
	return value
		.replace(/&/g, "&amp;")
		.replace(/</g, "&lt;")
		.replace(/>/g, "&gt;")
		.replace(/"/g, "&quot;");
}

export { PRESENTATION, P_NS };
