import type { ColorContext, Theme } from "./color";
import { gradientCss, resolveColor, resolveFillColor } from "./color";
import type { PptxPackage } from "./package";
import type { Crop, Fill, Stroke } from "./types";
import { emuToPx } from "./types";
import { attr, child, numAttr } from "./xml";

/**
 * Everything a shape needs in order to resolve references: the package (for
 * media), the path of the part it lives in (for relationships) and the theme.
 */
export interface ParseContext {
	pkg: PptxPackage;
	partPath: string;
	theme: Theme;
}

export function colorContext(ctx: ParseContext, phClr?: string): ColorContext {
	return { theme: ctx.theme, phClr };
}

/**
 * Parse the fill declared directly on a properties element (a:spPr, a:tcPr,
 * p:bgPr, ...). Returns null when no fill is declared at all, which means
 * "inherit"; a `none` kind means an explicit a:noFill.
 */
export function parseFill(props: Element | null, ctx: ParseContext, phClr?: string): Fill | null {
	if (!props) return null;
	const cc = colorContext(ctx, phClr);
	for (let n = props.firstElementChild; n; n = n.nextElementSibling) {
		switch (n.localName) {
			case "noFill":
				return { kind: "none" };
			case "solidFill": {
				const color = resolveFillColor(n, cc);
				return color ? { kind: "solid", color } : null;
			}
			case "gradFill": {
				const css = gradientCss(n, cc);
				return css ? { kind: "gradient", css } : null;
			}
			case "blipFill":
				return parseBlipFill(n, ctx);
			case "pattFill": {
				// Approximate a pattern with its foreground colour; a flat fill reads
				// far better than dropping the shape's colour entirely.
				const fg = resolveFillColor(child(n, "fgClr"), cc);
				const bg = resolveFillColor(child(n, "bgClr"), cc);
				const color = fg ?? bg;
				return color ? { kind: "solid", color } : null;
			}
			case "grpFill":
				return null;
		}
	}
	return null;
}

function parseBlipFill(blipFill: Element, ctx: ParseContext): Fill {
	const blip = child(blipFill, "blip");
	const path = ctx.pkg.relTarget(ctx.partPath, attr(blip, "embed"));
	const alphaMod = child(blip, "alphaModFix");
	const opacity = alphaMod ? (numAttr(alphaMod, "amt") ?? 100000) / 100000 : 1;
	return {
		kind: "image",
		url: ctx.pkg.mediaUrl(path),
		mediaPath: path,
		opacity,
	};
}

/** a:srcRect crop values, as fractions of the source image. */
export function parseCrop(blipFill: Element | null): Crop | null {
	const src = child(blipFill, "srcRect");
	if (!src) return null;
	const l = (numAttr(src, "l") ?? 0) / 100000;
	const t = (numAttr(src, "t") ?? 0) / 100000;
	const r = (numAttr(src, "r") ?? 0) / 100000;
	const b = (numAttr(src, "b") ?? 0) / 100000;
	if (!l && !t && !r && !b) return null;
	return { l, t, r, b };
}

const DASH_STYLES: Record<string, string> = {
	solid: "solid",
	dot: "dotted",
	sysDot: "dotted",
	dash: "dashed",
	sysDash: "dashed",
	lgDash: "dashed",
	dashDot: "dashed",
	sysDashDot: "dashed",
	lgDashDot: "dashed",
	lgDashDotDot: "dashed",
	sysDashDotDot: "dashed",
};

/** Parse an a:ln outline. Returns null for "no outline" or "inherit". */
export function parseStroke(props: Element | null, ctx: ParseContext, phClr?: string): Stroke | null {
	const ln = child(props, "ln");
	if (!ln) return null;
	if (child(ln, "noFill")) return null;
	const color = resolveFillColor(child(ln, "solidFill"), colorContext(ctx, phClr));
	if (!color) return null;
	const widthEmu = numAttr(ln, "w");
	const dash = attr(child(ln, "prstDash"), "val");
	return {
		color,
		// PowerPoint's default hairline is 0.75pt; anything thinner disappears.
		width: widthEmu === null ? 1 : Math.max(0.75, emuToPx(widthEmu)),
		style: (dash && DASH_STYLES[dash]) || "solid",
	};
}

/**
 * Resolve a p:style reference (a:fillRef / a:lnRef) against the theme's format
 * scheme. This is how most default autoshapes get their accent colour.
 */
export function fillFromStyleRef(style: Element | null, ctx: ParseContext): Fill | null {
	const fillRef = child(style, "fillRef");
	if (!fillRef) return null;
	const idx = numAttr(fillRef, "idx") ?? 0;
	const phClr = resolveFillColor(fillRef, colorContext(ctx)) ?? undefined;
	const source = idx >= 1001 ? ctx.theme.fmt.bgFillStyles[idx - 1001] : ctx.theme.fmt.fillStyles[idx - 1];
	if (!source) return phClr ? { kind: "solid", color: phClr } : null;
	// The style entry is itself a fill element; wrap it so parseFill sees it as a child.
	return parseFillElement(source, ctx, phClr);
}

export function strokeFromStyleRef(style: Element | null, ctx: ParseContext): Stroke | null {
	const lnRef = child(style, "lnRef");
	if (!lnRef) return null;
	const idx = numAttr(lnRef, "idx") ?? 0;
	const phClr = resolveFillColor(lnRef, colorContext(ctx)) ?? undefined;
	const source = ctx.theme.fmt.lineStyles[idx - 1];
	if (!source) return null;
	if (child(source, "noFill")) return null;
	const color = resolveFillColor(child(source, "solidFill"), colorContext(ctx, phClr));
	if (!color) return null;
	const widthEmu = numAttr(source, "w");
	const dash = attr(child(source, "prstDash"), "val");
	return {
		color,
		width: widthEmu === null ? 1 : Math.max(0.75, emuToPx(widthEmu)),
		style: (dash && DASH_STYLES[dash]) || "solid",
	};
}

/** The text colour implied by a p:style/a:fontRef, if it declares one. */
export function colorFromFontRef(style: Element | null, ctx: ParseContext): string | null {
	const fontRef = child(style, "fontRef");
	if (!fontRef) return null;
	return resolveFillColor(fontRef, colorContext(ctx));
}

/** Treat a bare fill element (as found in fillStyleLst) as if it were a fill child. */
function parseFillElement(el: Element, ctx: ParseContext, phClr?: string): Fill | null {
	const cc = colorContext(ctx, phClr);
	switch (el.localName) {
		case "noFill":
			return { kind: "none" };
		case "solidFill": {
			const color = resolveFillColor(el, cc);
			return color ? { kind: "solid", color } : null;
		}
		case "gradFill": {
			const css = gradientCss(el, cc);
			return css ? { kind: "gradient", css } : null;
		}
		case "blipFill":
			return parseBlipFill(el, ctx);
		default:
			return null;
	}
}

/** Resolve the colour used by a run's a:solidFill, honouring theme colours. */
export function textColor(rPr: Element | null, ctx: ParseContext, phClr?: string): string | null {
	const solid = child(rPr, "solidFill");
	if (solid) return resolveFillColor(solid, colorContext(ctx, phClr));
	const highlight = child(rPr, "highlight");
	if (highlight && !solid) return null;
	return null;
}

/** a:highlight, rendered as a background behind the run. */
export function highlightColor(rPr: Element | null, ctx: ParseContext): string | null {
	const hl = child(rPr, "highlight");
	if (!hl) return null;
	for (let n = hl.firstElementChild; n; n = n.nextElementSibling) {
		const c = resolveColor(n, colorContext(ctx));
		if (c) return c;
	}
	return null;
}
