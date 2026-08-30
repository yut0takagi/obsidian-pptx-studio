import type { PptxPackage } from "../pptx/package";
import { attr } from "../pptx/xml";

const CONTENT_TYPES = "[Content_Types].xml";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

/**
 * `[Content_Types].xml` is the part that tells a consumer what every other part
 * is. Adding a picture or a slide without registering it there produces a file
 * that opens as "repair needed", so every part-creating edit goes through here.
 */
function typesRoot(pkg: PptxPackage): Element | null {
	return pkg.xml(CONTENT_TYPES)?.documentElement ?? null;
}

/** Register a default content type for a file extension, e.g. png -> image/png. */
export function ensureDefault(pkg: PptxPackage, extension: string, contentType: string): void {
	const root = typesRoot(pkg);
	if (!root) return;
	const ext = extension.toLowerCase().replace(/^\./, "");
	for (let n = root.firstElementChild; n; n = n.nextElementSibling) {
		if (n.localName === "Default" && (attr(n, "Extension") ?? "").toLowerCase() === ext) return;
	}
	const el = root.ownerDocument.createElementNS(CT_NS, "Default");
	el.setAttribute("Extension", ext);
	el.setAttribute("ContentType", contentType);
	// Defaults conventionally precede overrides; keep the file tidy.
	const firstOverride = Array.from(root.children).find((c) => c.localName === "Override") ?? null;
	root.insertBefore(el, firstOverride);
	pkg.markDirty(CONTENT_TYPES);
}

/** Register the content type of one specific part, e.g. a new slide. */
export function ensureOverride(pkg: PptxPackage, partName: string, contentType: string): void {
	const root = typesRoot(pkg);
	if (!root) return;
	const name = partName.startsWith("/") ? partName : `/${partName}`;
	for (let n = root.firstElementChild; n; n = n.nextElementSibling) {
		if (n.localName === "Override" && attr(n, "PartName") === name) return;
	}
	const el = root.ownerDocument.createElementNS(CT_NS, "Override");
	el.setAttribute("PartName", name);
	el.setAttribute("ContentType", contentType);
	root.appendChild(el);
	pkg.markDirty(CONTENT_TYPES);
}

/** Drop the override for a part that is being deleted. */
export function removeOverride(pkg: PptxPackage, partName: string): void {
	const root = typesRoot(pkg);
	if (!root) return;
	const name = partName.startsWith("/") ? partName : `/${partName}`;
	for (const el of Array.from(root.children)) {
		if (el.localName === "Override" && attr(el, "PartName") === name) {
			root.removeChild(el);
			pkg.markDirty(CONTENT_TYPES);
			return;
		}
	}
}

export const CONTENT_TYPES_PART = CONTENT_TYPES;

export const CT = {
	slide: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
	notesSlide: "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
	viewProps: "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml",
	image: {
		png: "image/png",
		jpg: "image/jpeg",
		jpeg: "image/jpeg",
		gif: "image/gif",
		bmp: "image/bmp",
		svg: "image/svg+xml",
		webp: "image/webp",
	} as Record<string, string>,
};
