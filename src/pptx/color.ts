import { attr, child, children, numAttr } from "./xml";

export interface Theme {
	/** Scheme slot -> "#rrggbb", keyed by dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink. */
	scheme: Record<string, string>;
	/** p:clrMap on the master: bg1/tx1/bg2/tx2 -> scheme slot. */
	clrMap: Record<string, string>;
	majorFont: string;
	minorFont: string;
	/** a:fmtScheme lists, kept as elements so style references can be resolved lazily. */
	fmt: {
		fillStyles: Element[];
		lineStyles: Element[];
		bgFillStyles: Element[];
	};
}

export const DEFAULT_THEME: Theme = {
	scheme: {
		dk1: "#000000",
		lt1: "#ffffff",
		dk2: "#44546a",
		lt2: "#e7e6e6",
		accent1: "#4472c4",
		accent2: "#ed7d31",
		accent3: "#a5a5a5",
		accent4: "#ffc000",
		accent5: "#5b9bd5",
		accent6: "#70ad47",
		hlink: "#0563c1",
		folHlink: "#954f72",
	},
	clrMap: { bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2" },
	majorFont: "Calibri Light",
	minorFont: "Calibri",
	fmt: { fillStyles: [], lineStyles: [], bgFillStyles: [] },
};

/** Parse a theme part (ppt/theme/themeN.xml) into a colour scheme and font pair. */
export function parseTheme(themeRoot: Element | null, clrMapEl: Element | null): Theme {
	const scheme: Record<string, string> = { ...DEFAULT_THEME.scheme };
	const elements = child(themeRoot, "themeElements");
	const clrScheme = child(elements, "clrScheme");
	if (clrScheme) {
		for (let n = clrScheme.firstElementChild; n; n = n.nextElementSibling) {
			const value = literalColor(n.firstElementChild);
			if (value) scheme[n.localName] = value;
		}
	}

	const fontScheme = child(elements, "fontScheme");
	const majorFont = attr(child(child(fontScheme, "majorFont"), "latin"), "typeface");
	const minorFont = attr(child(child(fontScheme, "minorFont"), "latin"), "typeface");

	const clrMap: Record<string, string> = { ...DEFAULT_THEME.clrMap };
	if (clrMapEl) {
		for (const slot of ["bg1", "tx1", "bg2", "tx2"]) {
			const mapped = attr(clrMapEl, slot);
			if (mapped) clrMap[slot] = mapped;
		}
	}

	const fmtScheme = child(elements, "fmtScheme");
	const fmt = {
		fillStyles: elementChildren(child(fmtScheme, "fillStyleLst")),
		lineStyles: elementChildren(child(fmtScheme, "lnStyleLst")),
		bgFillStyles: elementChildren(child(fmtScheme, "bgFillStyleLst")),
	};

	return {
		scheme,
		clrMap,
		majorFont: majorFont || DEFAULT_THEME.majorFont,
		minorFont: minorFont || DEFAULT_THEME.minorFont,
		fmt,
	};
}

function elementChildren(el: Element | null): Element[] {
	if (!el) return [];
	const out: Element[] = [];
	for (let n = el.firstElementChild; n; n = n.nextElementSibling) out.push(n);
	return out;
}

/** Colour lookup context: the theme plus the current placeholder colour, if any. */
export interface ColorContext {
	theme: Theme;
	/** Value substituted for <a:schemeClr val="phClr"/> inside style matrices. */
	phClr?: string;
}

/**
 * Resolve a colour child element (srgbClr / schemeClr / sysClr / ...) into a CSS
 * colour, applying any tint/shade/lumMod/lumOff/satMod/alpha transforms on it.
 */
export function resolveColor(el: Element | null | undefined, ctx: ColorContext): string | null {
	if (!el) return null;
	let hex = literalColor(el, ctx);
	if (!hex) return null;

	let rgb = hexToRgb(hex);
	let alpha = 1;

	for (let n = el.firstElementChild; n; n = n.nextElementSibling) {
		const val = numAttr(n, "val");
		switch (n.localName) {
			case "alpha":
				if (val !== null) alpha *= val / 100000;
				break;
			case "alphaOff":
				if (val !== null) alpha = clamp01(alpha + val / 100000);
				break;
			case "tint":
				if (val !== null) rgb = applyTint(rgb, val / 100000);
				break;
			case "shade":
				if (val !== null) rgb = applyShade(rgb, val / 100000);
				break;
			case "lumMod":
				if (val !== null) rgb = adjustLuminance(rgb, val / 100000, 0);
				break;
			case "lumOff":
				if (val !== null) rgb = adjustLuminance(rgb, 1, val / 100000);
				break;
			case "satMod":
				if (val !== null) rgb = adjustSaturation(rgb, val / 100000);
				break;
			case "gray":
				rgb = toGray(rgb);
				break;
			case "inv":
				rgb = [255 - rgb[0], 255 - rgb[1], 255 - rgb[2]];
				break;
			case "comp":
			case "invGamma":
			case "gamma":
				// Rare enough that approximating with a no-op beats guessing wrong.
				break;
		}
	}

	hex = rgbToHex(rgb);
	return alpha >= 0.999 ? hex : `rgba(${rgb[0]}, ${rgb[1]}, ${rgb[2]}, ${round(alpha, 3)})`;
}

/** The base colour of a colour element, before transforms. */
function literalColor(el: Element | null | undefined, ctx?: ColorContext): string | null {
	if (!el) return null;
	switch (el.localName) {
		case "srgbClr": {
			const v = attr(el, "val");
			return v ? normaliseHex(v) : null;
		}
		case "sysClr": {
			const last = attr(el, "lastClr");
			if (last) return normaliseHex(last);
			return attr(el, "val") === "window" ? "#ffffff" : "#000000";
		}
		case "schemeClr": {
			const v = attr(el, "val");
			if (!v) return null;
			if (v === "phClr") return ctx?.phClr ?? null;
			const theme = ctx?.theme ?? DEFAULT_THEME;
			const slot = theme.clrMap[v] ?? v;
			// "bg1"/"tx1" map through clrMap; dk1/lt1 also answer to "dk1"/"lt1".
			return theme.scheme[slot] ?? theme.scheme[v] ?? null;
		}
		case "prstClr": {
			const v = attr(el, "val");
			return v ? (PRESET_COLORS[v] ?? null) : null;
		}
		case "scrgbClr": {
			const r = numAttr(el, "r") ?? 0;
			const g = numAttr(el, "g") ?? 0;
			const b = numAttr(el, "b") ?? 0;
			return rgbToHex([
				Math.round((r / 100000) * 255),
				Math.round((g / 100000) * 255),
				Math.round((b / 100000) * 255),
			]);
		}
		case "hslClr": {
			const h = (numAttr(el, "hue") ?? 0) / 60000;
			const s = (numAttr(el, "sat") ?? 0) / 100000;
			const l = (numAttr(el, "lum") ?? 0) / 100000;
			return rgbToHex(hslToRgb([h, s, l]));
		}
		default:
			return null;
	}
}

/** Find and resolve the colour inside a fill-like wrapper (a:solidFill, a:gs, ...). */
export function resolveFillColor(wrapper: Element | null | undefined, ctx: ColorContext): string | null {
	if (!wrapper) return null;
	for (let n = wrapper.firstElementChild; n; n = n.nextElementSibling) {
		const c = resolveColor(n, ctx);
		if (c) return c;
	}
	return null;
}

/** Build a CSS linear-gradient() from an a:gradFill. */
export function gradientCss(gradFill: Element, ctx: ColorContext): string | null {
	const stops = children(child(gradFill, "gsLst"), "gs")
		.map((gs) => {
			const pos = (numAttr(gs, "pos") ?? 0) / 1000;
			const color = resolveFillColor(gs, ctx);
			return color ? { pos, color } : null;
		})
		.filter((s): s is { pos: number; color: string } => s !== null)
		.sort((a, b) => a.pos - b.pos);
	if (stops.length === 0) return null;
	if (stops.length === 1) return `linear-gradient(${stops[0].color}, ${stops[0].color})`;

	const path = child(gradFill, "path");
	const stopList = stops.map((s) => `${s.color} ${round(s.pos, 2)}%`).join(", ");
	if (path) {
		const shape = attr(path, "path") === "circle" ? "circle" : "ellipse";
		return `radial-gradient(${shape} at 50% 50%, ${stopList})`;
	}
	// OOXML angles run clockwise from "pointing right"; CSS runs clockwise from
	// "pointing up", so the same direction is the OOXML angle plus 90 degrees.
	const lin = child(gradFill, "lin");
	const angle = ((numAttr(lin, "ang") ?? 5400000) / 60000 + 90) % 360;
	return `linear-gradient(${round(angle, 1)}deg, ${stopList})`;
}

// ---------------------------------------------------------------- colour math

type Rgb = [number, number, number];

function normaliseHex(v: string): string {
	const clean = v.replace(/^#/, "").trim();
	if (/^[0-9a-fA-F]{6}$/.test(clean)) return `#${clean.toLowerCase()}`;
	if (/^[0-9a-fA-F]{3}$/.test(clean)) {
		return `#${clean
			.toLowerCase()
			.split("")
			.map((c) => c + c)
			.join("")}`;
	}
	return "#000000";
}

function hexToRgb(hex: string): Rgb {
	const h = normaliseHex(hex).slice(1);
	return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
}

function rgbToHex(rgb: Rgb): string {
	return (
		"#" +
		rgb
			.map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0"))
			.join("")
	);
}

function applyTint(rgb: Rgb, tint: number): Rgb {
	return rgb.map((c) => c * tint + 255 * (1 - tint)) as Rgb;
}

function applyShade(rgb: Rgb, shade: number): Rgb {
	return rgb.map((c) => c * shade) as Rgb;
}

function adjustLuminance(rgb: Rgb, mod: number, off: number): Rgb {
	const [h, s, l] = rgbToHsl(rgb);
	return hslToRgb([h, s, clamp01(l * mod + off)]);
}

function adjustSaturation(rgb: Rgb, mod: number): Rgb {
	const [h, s, l] = rgbToHsl(rgb);
	return hslToRgb([h, clamp01(s * mod), l]);
}

function toGray(rgb: Rgb): Rgb {
	const y = 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2];
	return [y, y, y];
}

/** Returns hue in degrees, saturation and lightness in 0..1. */
function rgbToHsl(rgb: Rgb): [number, number, number] {
	const [r, g, b] = rgb.map((c) => c / 255) as Rgb;
	const max = Math.max(r, g, b);
	const min = Math.min(r, g, b);
	const l = (max + min) / 2;
	if (max === min) return [0, 0, l];
	const d = max - min;
	const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
	let h: number;
	if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) * 60;
	else if (max === g) h = ((b - r) / d + 2) * 60;
	else h = ((r - g) / d + 4) * 60;
	return [h, s, l];
}

function hslToRgb([h, s, l]: [number, number, number]): Rgb {
	if (s === 0) {
		const v = l * 255;
		return [v, v, v];
	}
	const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
	const p = 2 * l - q;
	const hk = (((h % 360) + 360) % 360) / 360;
	const channel = (t: number): number => {
		let tt = t;
		if (tt < 0) tt += 1;
		if (tt > 1) tt -= 1;
		if (tt < 1 / 6) return p + (q - p) * 6 * tt;
		if (tt < 1 / 2) return q;
		if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6;
		return p;
	};
	return [channel(hk + 1 / 3) * 255, channel(hk) * 255, channel(hk - 1 / 3) * 255];
}

function clamp01(v: number): number {
	return Math.max(0, Math.min(1, v));
}

function round(v: number, digits: number): number {
	const f = 10 ** digits;
	return Math.round(v * f) / f;
}

/** The subset of ECMA-376 preset colours that actually shows up in decks. */
const PRESET_COLORS: Record<string, string> = {
	black: "#000000",
	white: "#ffffff",
	red: "#ff0000",
	green: "#008000",
	blue: "#0000ff",
	yellow: "#ffff00",
	cyan: "#00ffff",
	aqua: "#00ffff",
	magenta: "#ff00ff",
	fuchsia: "#ff00ff",
	gray: "#808080",
	grey: "#808080",
	darkGray: "#a9a9a9",
	lightGray: "#d3d3d3",
	silver: "#c0c0c0",
	maroon: "#800000",
	olive: "#808000",
	purple: "#800080",
	teal: "#008080",
	navy: "#000080",
	lime: "#00ff00",
	orange: "#ffa500",
	pink: "#ffc0cb",
	brown: "#a52a2a",
	gold: "#ffd700",
	indigo: "#4b0082",
	violet: "#ee82ee",
	beige: "#f5f5dc",
	ivory: "#fffff0",
	khaki: "#f0e68c",
	lavender: "#e6e6fa",
	salmon: "#fa8072",
	tan: "#d2b48c",
	turquoise: "#40e0d0",
	wheat: "#f5deb3",
	darkBlue: "#00008b",
	darkGreen: "#006400",
	darkRed: "#8b0000",
	lightBlue: "#add8e6",
	lightGreen: "#90ee90",
	dkGray: "#a9a9a9",
	ltGray: "#d3d3d3",
};
