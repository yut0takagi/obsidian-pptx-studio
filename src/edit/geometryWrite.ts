import type { Frame, Shape } from "../pptx/types";
import { EMU_PER_PX } from "../pptx/types";
import { child } from "../pptx/xml";

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";

const emu = (px: number) => Math.round(px * EMU_PER_PX);

/** The element that owns a shape's a:xfrm, and where the xfrm belongs inside it. */
function transformHost(source: Element): { host: Element; ns: string; after: string | null } | null {
	switch (source.localName) {
		case "sp":
		case "pic":
		case "cxnSp": {
			const spPr = child(source, "spPr");
			return spPr ? { host: spPr, ns: A_NS, after: null } : null;
		}
		case "grpSp": {
			const grpSpPr = child(source, "grpSpPr");
			return grpSpPr ? { host: grpSpPr, ns: A_NS, after: null } : null;
		}
		case "graphicFrame":
			// A graphic frame carries p:xfrm directly, between nvGraphicFramePr and graphic.
			return { host: source, ns: P_NS, after: "nvGraphicFramePr" };
		default:
			return null;
	}
}

/**
 * Write a shape's position and size back into its XML.
 *
 * A placeholder that inherited its frame from the layout has no a:xfrm of its
 * own; moving it creates one, which is exactly what PowerPoint does.
 */
export function writeShapeFrame(shape: Shape, frame: Frame): string | null {
	const source = shape.source;
	if (!source) return null;
	const target = transformHost(source);
	if (!target) return null;

	const doc = source.ownerDocument;
	const prefix = target.ns === P_NS ? "p" : "a";
	let xfrm = child(target.host, "xfrm");
	if (!xfrm) {
		xfrm = doc.createElementNS(target.ns, `${prefix}:xfrm`);
		const anchor = target.after ? child(target.host, target.after) : null;
		target.host.insertBefore(xfrm, anchor ? anchor.nextSibling : target.host.firstChild);
	}

	let off = child(xfrm, "off");
	if (!off) {
		off = doc.createElementNS(A_NS, "a:off");
		xfrm.insertBefore(off, xfrm.firstChild);
	}
	let ext = child(xfrm, "ext");
	if (!ext) {
		ext = doc.createElementNS(A_NS, "a:ext");
		off.parentNode?.insertBefore(ext, off.nextSibling);
	}

	off.setAttribute("x", String(emu(frame.x)));
	off.setAttribute("y", String(emu(frame.y)));
	ext.setAttribute("cx", String(emu(Math.max(frame.w, 0))));
	ext.setAttribute("cy", String(emu(Math.max(frame.h, 0))));

	return shape.sourcePart;
}
