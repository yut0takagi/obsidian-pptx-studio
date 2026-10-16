import type { PptxPackage, Relationship } from "../pptx/package";
import { attr, parseXml } from "../pptx/xml";

export const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
export const REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** The `_rels/<name>.rels` path for a part. */
export function relsPathFor(partPath: string): string {
	const slash = partPath.lastIndexOf("/");
	const dir = slash === -1 ? "" : partPath.slice(0, slash);
	const base = slash === -1 ? partPath : partPath.slice(slash + 1);
	return `${dir ? `${dir}/` : ""}_rels/${base}.rels`;
}

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const EMPTY_RELS = `${DECL}<Relationships xmlns="${PACKAGE_REL_NS}"/>`;

function relsDocument(pkg: PptxPackage, partPath: string): { path: string; doc: Document } {
	const path = relsPathFor(partPath);
	let doc = pkg.xml(path);
	if (!doc?.documentElement) {
		pkg.replacePart(path, new TextEncoder().encode(EMPTY_RELS));
		doc = pkg.xml(path);
	}
	if (!doc?.documentElement) throw new Error(`Could not create relationships for ${partPath}`);
	return { path, doc };
}

/** Express a package-absolute target relative to the part that references it. */
export function relativeTarget(fromPart: string, target: string): string {
	const from = fromPart.split("/").slice(0, -1);
	const to = target.split("/");
	let i = 0;
	while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
	const up = from.slice(i).map(() => "..");
	return [...up, ...to.slice(i)].join("/");
}

/**
 * Find or create a relationship from `partPath` to `target`, returning its r:id.
 *
 * Reusing an existing relationship matters: pasting the same picture ten times
 * should add one relationship and ten references to it, not ten relationships
 * to the same image.
 */
export function ensureRelationship(
	pkg: PptxPackage,
	partPath: string,
	type: string,
	target: string,
	external = false,
): string {
	const relative = external ? target : relativeTarget(partPath, target);
	const { path, doc } = relsDocument(pkg, partPath);
	const root = doc.documentElement;

	let maxId = 0;
	for (const el of Array.from(root.getElementsByTagName("*"))) {
		if (el.localName !== "Relationship") continue;
		const id = attr(el, "Id") ?? "";
		const n = Number(/^rId(\d+)$/.exec(id)?.[1] ?? 0);
		if (n > maxId) maxId = n;
		if (
			attr(el, "Type") === type &&
			attr(el, "Target") === relative &&
			(attr(el, "TargetMode") === "External") === external
		) {
			return id;
		}
	}

	const id = `rId${maxId + 1}`;
	const rel = doc.createElementNS(PACKAGE_REL_NS, "Relationship");
	rel.setAttribute("Id", id);
	rel.setAttribute("Type", type);
	rel.setAttribute("Target", relative);
	if (external) rel.setAttribute("TargetMode", "External");
	root.appendChild(rel);

	pkg.markDirty(path);
	pkg.invalidateRels();
	return id;
}

/** Remove a relationship by id, if present. */
export function removeRelationship(pkg: PptxPackage, partPath: string, id: string): void {
	const path = relsPathFor(partPath);
	const doc = pkg.xml(path);
	if (!doc?.documentElement) return;
	for (const el of Array.from(doc.documentElement.getElementsByTagName("*"))) {
		if (el.localName === "Relationship" && attr(el, "Id") === id) {
			el.parentNode?.removeChild(el);
			pkg.markDirty(path);
			pkg.invalidateRels();
			return;
		}
	}
}

/** Point an existing relationship at a different part, keeping its id. */
export function retargetRelationship(
	pkg: PptxPackage,
	partPath: string,
	id: string,
	target: string,
): boolean {
	const path = relsPathFor(partPath);
	const doc = pkg.xml(path);
	if (!doc?.documentElement) return false;
	for (const el of Array.from(doc.documentElement.getElementsByTagName("*"))) {
		if (el.localName !== "Relationship" || attr(el, "Id") !== id) continue;
		el.setAttribute("Target", relativeTarget(partPath, target));
		pkg.markDirty(path);
		pkg.invalidateRels();
		return true;
	}
	return false;
}

/**
 * Give `toPart` the relationships `fromPart` has, with the same ids and types,
 * each target re-expressed relative to where `toPart` lives. `retarget` can send
 * one somewhere else — a copied notes slide pointing at the copied slide rather
 * than at the original.
 */
export function copyRelationships(
	pkg: PptxPackage,
	fromPart: string,
	toPart: string,
	retarget: (rel: Relationship) => string = (rel) => rel.target,
): void {
	const doc = parseXml(emptyRelsBytes());
	const root = doc.documentElement;
	for (const rel of pkg.rels(fromPart).values()) {
		const el = doc.createElementNS(PACKAGE_REL_NS, "Relationship");
		el.setAttribute("Id", rel.id);
		el.setAttribute("Type", rel.type);
		el.setAttribute("Target", rel.external ? rel.target : relativeTarget(toPart, retarget(rel)));
		if (rel.external) el.setAttribute("TargetMode", "External");
		root.appendChild(el);
	}
	const body = new XMLSerializer().serializeToString(doc).replace(/^\s*<\?xml[^?]*\?>\s*/, "");
	pkg.replacePart(relsPathFor(toPart), new TextEncoder().encode(DECL + body));
}

/** The rels parts that must be snapshotted when relationships may change. */
export function relsParts(...partPaths: string[]): string[] {
	return partPaths.map(relsPathFor);
}

/** Parse a standalone rels document, used when building brand new parts. */
export function emptyRelsBytes(): Uint8Array {
	return new TextEncoder().encode(EMPTY_RELS);
}

/** Guard against a malformed rels part silently swallowing edits. */
export function assertRelsParsable(bytes: Uint8Array): void {
	parseXml(bytes);
}
