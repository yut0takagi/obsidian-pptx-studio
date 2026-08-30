import type { PptxPackage } from "../pptx/package";
import { child, children } from "../pptx/xml";
import { REL_NS, ensureRelationship } from "./rels";

export const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";

/** Element names that count as a shape in a slide's shape tree. */
const SHAPE_NAMES = new Set(["sp", "pic", "graphicFrame", "grpSp", "cxnSp", "AlternateContent"]);

/** A slide's p:spTree, the container every shape lives in. */
export function spTreeOf(pkg: PptxPackage, slidePath: string): Element | null {
	const root = pkg.xml(slidePath)?.documentElement ?? null;
	return child(child(root, "cSld"), "spTree");
}

/** Top-level shapes, in z-order: first element is furthest back. */
export function shapeElements(spTree: Element | null): Element[] {
	if (!spTree) return [];
	const out: Element[] = [];
	for (let n = spTree.firstElementChild; n; n = n.nextElementSibling) {
		if (SHAPE_NAMES.has(n.localName)) out.push(n);
	}
	return out;
}

/** The non-visual properties element that carries a shape's id and name. */
export function nonVisualProps(el: Element): Element | null {
	for (const container of ["nvSpPr", "nvPicPr", "nvGrpSpPr", "nvCxnSpPr", "nvGraphicFramePr"]) {
		const found = child(child(el, container), "cNvPr");
		if (found) return found;
	}
	return null;
}

export function shapeIdOf(el: Element): string | null {
	return nonVisualProps(el)?.getAttribute("id") ?? null;
}

/** One past the largest shape id in a slide, so new shapes cannot collide. */
export function nextShapeId(spTree: Element): number {
	let max = 1;
	for (const el of Array.from(spTree.getElementsByTagName("*"))) {
		if (el.localName !== "cNvPr") continue;
		const n = Number(el.getAttribute("id") ?? 0);
		if (Number.isFinite(n) && n > max) max = n;
	}
	return max + 1;
}

/** Give every shape in a subtree a fresh id, so a paste cannot duplicate one. */
export function reassignIds(el: Element, next: () => number): void {
	const targets = el.localName === "cNvPr" ? [el] : [];
	for (const node of Array.from(el.getElementsByTagName("*"))) {
		if (node.localName === "cNvPr") targets.push(node);
	}
	for (const node of targets) node.setAttribute("id", String(next()));
}

/**
 * Rewrite the relationship ids inside a subtree so they resolve against a
 * different part.
 *
 * Without this, pasting a picture onto another slide silently produces a shape
 * pointing at an r:id that slide has never heard of, and the image vanishes.
 */
export function remapRelationships(
	pkg: PptxPackage,
	el: Element,
	fromPart: string,
	toPart: string,
): void {
	if (fromPart === toPart) return;
	const source = pkg.rels(fromPart);
	const nodes = [el, ...Array.from(el.getElementsByTagName("*"))];
	for (const node of nodes) {
		for (const a of Array.from(node.attributes)) {
			if (a.namespaceURI !== REL_NS) continue;
			const rel = source.get(a.value);
			if (!rel) continue;
			node.setAttributeNS(
				REL_NS,
				a.name,
				ensureRelationship(pkg, toPart, rel.type, rel.target, rel.external),
			);
		}
	}
}

/** Create an element in the DrawingML namespace. */
export function a(doc: Document, name: string): Element {
	return doc.createElementNS(A_NS, `a:${name}`);
}

/** Create an element in the PresentationML namespace. */
export function p(doc: Document, name: string): Element {
	return doc.createElementNS(P_NS, `p:${name}`);
}

/** Insert `el` so that it sits immediately after `reference` in z-order. */
export function insertAfter(spTree: Element, el: Element, reference: Element | null): void {
	spTree.insertBefore(el, reference ? reference.nextSibling : null);
}

/** Remove a shape from its tree. */
export function removeShape(el: Element): void {
	el.parentNode?.removeChild(el);
}

/** Set the visible name shown in selection panes and alt text. */
export function setShapeName(el: Element, name: string): void {
	nonVisualProps(el)?.setAttribute("name", name);
}

/** The shape elements in a slide matching a set of ids, in z-order. */
export function findShapesById(spTree: Element | null, ids: Set<string>): Element[] {
	return shapeElements(spTree).filter((el) => {
		const id = shapeIdOf(el);
		return id !== null && ids.has(id);
	});
}

/** Direct children of a group that are themselves shapes. */
export function groupChildren(group: Element): Element[] {
	const out: Element[] = [];
	for (let n = group.firstElementChild; n; n = n.nextElementSibling) {
		if (SHAPE_NAMES.has(n.localName)) out.push(n);
	}
	return out;
}

/** a:ext / a:off values on a group's child coordinate space. */
export function groupChildSpace(group: Element): Element | null {
	return child(child(group, "grpSpPr"), "xfrm");
}

/** Convenience for reading a list of direct children by name. */
export { child, children };
