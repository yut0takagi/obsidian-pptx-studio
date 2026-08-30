import type { Frame } from "./types";
import { emuToPx } from "./types";
import { attr, boolAttr, child, numAttr } from "./xml";

export const EMPTY_FRAME: Frame = { x: 0, y: 0, w: 0, h: 0, rot: 0, flipH: false, flipV: false };

/** Read an a:xfrm into a pixel frame. Returns null when the element is absent. */
export function parseFrame(xfrm: Element | null): Frame | null {
	if (!xfrm) return null;
	const off = child(xfrm, "off");
	const ext = child(xfrm, "ext");
	return {
		x: emuToPx(numAttr(off, "x") ?? 0),
		y: emuToPx(numAttr(off, "y") ?? 0),
		w: emuToPx(numAttr(ext, "cx") ?? 0),
		h: emuToPx(numAttr(ext, "cy") ?? 0),
		rot: (numAttr(xfrm, "rot") ?? 0) / 60000,
		flipH: boolAttr(xfrm, "flipH") ?? false,
		flipV: boolAttr(xfrm, "flipV") ?? false,
	};
}

/** The child coordinate space declared by a group (a:chOff / a:chExt). */
export function parseChildFrame(xfrm: Element | null): { x: number; y: number; w: number; h: number } | null {
	if (!xfrm) return null;
	const off = child(xfrm, "chOff");
	const ext = child(xfrm, "chExt");
	if (!off && !ext) return null;
	return {
		x: emuToPx(numAttr(off, "x") ?? 0),
		y: emuToPx(numAttr(off, "y") ?? 0),
		w: emuToPx(numAttr(ext, "cx") ?? 0),
		h: emuToPx(numAttr(ext, "cy") ?? 0),
	};
}

/** The preset name on a:prstGeom, or "custGeom"/"rect" as a fallback. */
export function geometryName(spPr: Element | null): string {
	const prst = attr(child(spPr, "prstGeom"), "prst");
	if (prst) return prst;
	if (child(spPr, "custGeom")) return "custGeom";
	return "rect";
}

export interface GeometryCss {
	borderRadius?: string;
	clipPath?: string;
	/** True when the outline must be drawn as a background rather than a CSS border. */
	clipped: boolean;
}

/**
 * Map a preset shape to CSS. Anything unmapped falls back to a rectangle, which
 * keeps position and fill correct even when the silhouette is not.
 */
export function geometryCss(name: string, w: number, h: number): GeometryCss {
	switch (name) {
		case "rect":
		case "custGeom":
			return { clipped: false };
		case "roundRect":
		case "round1Rect":
		case "round2SameRect":
		case "round2DiagRect":
			return { borderRadius: `${Math.min(w, h) * 0.16}px`, clipped: false };
		case "ellipse":
		case "circle":
			return { borderRadius: "50%", clipped: false };
		case "pie":
		case "chord":
		case "blockArc":
			return { borderRadius: "50%", clipped: false };
		case "triangle":
			return { clipPath: "polygon(50% 0%, 100% 100%, 0% 100%)", clipped: true };
		case "rtTriangle":
			return { clipPath: "polygon(0% 0%, 0% 100%, 100% 100%)", clipped: true };
		case "diamond":
			return { clipPath: "polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)", clipped: true };
		case "parallelogram":
			return { clipPath: "polygon(20% 0%, 100% 0%, 80% 100%, 0% 100%)", clipped: true };
		case "trapezoid":
			return { clipPath: "polygon(20% 0%, 80% 0%, 100% 100%, 0% 100%)", clipped: true };
		case "pentagon":
			return {
				clipPath: "polygon(50% 0%, 100% 38%, 82% 100%, 18% 100%, 0% 38%)",
				clipped: true,
			};
		case "hexagon":
			return {
				clipPath: "polygon(25% 0%, 75% 0%, 100% 50%, 75% 100%, 25% 100%, 0% 50%)",
				clipped: true,
			};
		case "octagon":
			return {
				clipPath:
					"polygon(30% 0%, 70% 0%, 100% 30%, 100% 70%, 70% 100%, 30% 100%, 0% 70%, 0% 30%)",
				clipped: true,
			};
		case "star5":
			return {
				clipPath:
					"polygon(50% 0%, 61% 35%, 98% 35%, 68% 57%, 79% 91%, 50% 70%, 21% 91%, 32% 57%, 2% 35%, 39% 35%)",
				clipped: true,
			};
		case "star4":
			return {
				clipPath:
					"polygon(50% 0%, 62% 38%, 100% 50%, 62% 62%, 50% 100%, 38% 62%, 0% 50%, 38% 38%)",
				clipped: true,
			};
		case "rightArrow":
			return {
				clipPath: "polygon(0% 30%, 60% 30%, 60% 0%, 100% 50%, 60% 100%, 60% 70%, 0% 70%)",
				clipped: true,
			};
		case "leftArrow":
			return {
				clipPath: "polygon(100% 30%, 40% 30%, 40% 0%, 0% 50%, 40% 100%, 40% 70%, 100% 70%)",
				clipped: true,
			};
		case "upArrow":
			return {
				clipPath: "polygon(30% 100%, 30% 40%, 0% 40%, 50% 0%, 100% 40%, 70% 40%, 70% 100%)",
				clipped: true,
			};
		case "downArrow":
			return {
				clipPath: "polygon(30% 0%, 30% 60%, 0% 60%, 50% 100%, 100% 60%, 70% 60%, 70% 0%)",
				clipped: true,
			};
		case "chevron":
			return {
				clipPath: "polygon(0% 0%, 75% 0%, 100% 50%, 75% 100%, 0% 100%, 25% 50%)",
				clipped: true,
			};
		case "homePlate":
			return { clipPath: "polygon(0% 0%, 75% 0%, 100% 50%, 75% 100%, 0% 100%)", clipped: true };
		case "plus":
		case "mathPlus":
			return {
				clipPath:
					"polygon(35% 0%, 65% 0%, 65% 35%, 100% 35%, 100% 65%, 65% 65%, 65% 100%, 35% 100%, 35% 65%, 0% 65%, 0% 35%, 35% 35%)",
				clipped: true,
			};
		case "flowChartDecision":
			return { clipPath: "polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)", clipped: true };
		case "flowChartTerminator":
			return { borderRadius: `${h / 2}px`, clipped: false };
		case "flowChartProcess":
		case "flowChartPredefinedProcess":
			return { clipped: false };
		case "cloud":
			return { borderRadius: "50% 45% 55% 40% / 60% 55% 45% 50%", clipped: false };
		case "wedgeRectCallout":
		case "wedgeRoundRectCallout":
			return { borderRadius: `${Math.min(w, h) * 0.12}px`, clipped: false };
		case "line":
		case "straightConnector1":
			return { clipped: false };
		default:
			return { clipped: false };
	}
}

/** Preset names that are really lines, so we render a stroke rather than a box. */
export function isLineGeometry(name: string): boolean {
	return (
		name === "line" ||
		name === "straightConnector1" ||
		name.startsWith("bentConnector") ||
		name.startsWith("curvedConnector")
	);
}
