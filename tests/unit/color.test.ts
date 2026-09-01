/**
 * Theme parsing and colour resolution.
 *
 * Almost nothing in a deck states its colour outright: it names a theme slot,
 * which the master's colour map may have redirected, and then modifies it with
 * a transform expressed in hundred-thousandths. Getting any step of that wrong
 * shows up as a slide that is the wrong colour but not obviously broken, which
 * is exactly the kind of thing a rendering check does not catch.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import { installDomParser } from "./dom";

installDomParser();

import {
	type ColorContext,
	DEFAULT_THEME,
	gradientCss,
	parseTheme,
	resolveColor,
	resolveFillColor,
} from "../../src/pptx/color";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P = "http://schemas.openxmlformats.org/presentationml/2006/main";

function el(xml: string): Element {
	return new DOMParser().parseFromString(xml, "application/xml").documentElement;
}

/** A colour element, namespaced, without repeating the xmlns in every case. */
function color(tag: string, attrs: string, inner = ""): Element {
	return el(`<a:${tag} xmlns:a="${A}" ${attrs}>${inner}</a:${tag}>`);
}

const ctx: ColorContext = { theme: DEFAULT_THEME };

describe("literal colours", () => {
	it("normalises a six-digit hex to lower case", () => {
		assert.equal(resolveColor(color("srgbClr", `val="FF0000"`), ctx), "#ff0000");
	});

	it("expands a three-digit hex", () => {
		assert.equal(resolveColor(color("srgbClr", `val="abc"`), ctx), "#aabbcc");
	});

	it("falls back to black rather than emitting an invalid CSS colour", () => {
		assert.equal(resolveColor(color("srgbClr", `val="not-a-colour"`), ctx), "#000000");
	});

	it("prefers the lastClr PowerPoint cached over guessing at a system colour", () => {
		assert.equal(resolveColor(color("sysClr", `val="windowText" lastClr="C0C0C0"`), ctx), "#c0c0c0");
	});

	it("falls back to window white and windowText black with no cached value", () => {
		assert.equal(resolveColor(color("sysClr", `val="window"`), ctx), "#ffffff");
		assert.equal(resolveColor(color("sysClr", `val="windowText"`), ctx), "#000000");
	});

	it("reads a preset colour by name", () => {
		assert.equal(resolveColor(color("prstClr", `val="red"`), ctx), "#ff0000");
	});

	it("returns null for a preset outside the table, so a caller can fall back", () => {
		assert.equal(resolveColor(color("prstClr", `val="chartreuse"`), ctx), null);
	});

	it("reads scRGB percentages", () => {
		assert.equal(resolveColor(color("scrgbClr", `r="100000" g="0" b="0"`), ctx), "#ff0000");
	});

	it("reads HSL in its OOXML units", () => {
		assert.equal(resolveColor(color("hslClr", `hue="0" sat="100000" lum="50000"`), ctx), "#ff0000");
	});

	it("returns null for an absent or unrecognised element", () => {
		assert.equal(resolveColor(null, ctx), null);
		assert.equal(resolveColor(color("bogusClr", `val="FF0000"`), ctx), null);
	});
});

describe("scheme colours", () => {
	it("maps bg1 and tx1 through the colour map", () => {
		assert.equal(resolveColor(color("schemeClr", `val="tx1"`), ctx), DEFAULT_THEME.scheme.dk1);
		assert.equal(resolveColor(color("schemeClr", `val="bg1"`), ctx), DEFAULT_THEME.scheme.lt1);
	});

	it("answers to a slot named directly, which the colour map does not cover", () => {
		assert.equal(resolveColor(color("schemeClr", `val="accent1"`), ctx), "#4472c4");
		assert.equal(resolveColor(color("schemeClr", `val="dk1"`), ctx), "#000000");
	});

	it("substitutes phClr from the style matrix context", () => {
		assert.equal(
			resolveColor(color("schemeClr", `val="phClr"`), { theme: DEFAULT_THEME, phClr: "#123456" }),
			"#123456",
		);
	});

	it("returns null for phClr outside a style matrix, rather than inventing a colour", () => {
		assert.equal(resolveColor(color("schemeClr", `val="phClr"`), ctx), null);
	});
});

describe("colour transforms", () => {
	it("tints towards white and shades towards black", () => {
		assert.equal(
			resolveColor(color("srgbClr", `val="FF0000"`, `<a:tint val="50000"/>`), ctx),
			"#ff8080",
		);
		assert.equal(
			resolveColor(color("srgbClr", `val="FF0000"`, `<a:shade val="50000"/>`), ctx),
			"#800000",
		);
	});

	it("scales luminance with lumMod and offsets it with lumOff", () => {
		assert.equal(
			resolveColor(color("srgbClr", `val="FF0000"`, `<a:lumMod val="50000"/>`), ctx),
			"#800000",
		);
		assert.equal(
			resolveColor(color("srgbClr", `val="000000"`, `<a:lumOff val="20000"/>`), ctx),
			"#333333",
		);
	});

	it("desaturates with satMod", () => {
		assert.equal(
			resolveColor(color("srgbClr", `val="FF0000"`, `<a:satMod val="0"/>`), ctx),
			"#808080",
		);
	});

	it("inverts and greys", () => {
		assert.equal(resolveColor(color("srgbClr", `val="FF0000"`, `<a:inv/>`), ctx), "#00ffff");
		assert.equal(resolveColor(color("srgbClr", `val="FF0000"`, `<a:gray/>`), ctx), "#4c4c4c");
	});

	it("applies the transforms in the order they appear", () => {
		const el1 = color("srgbClr", `val="FF0000"`, `<a:inv/><a:shade val="50000"/>`);
		const el2 = color("srgbClr", `val="FF0000"`, `<a:shade val="50000"/><a:inv/>`);
		assert.equal(resolveColor(el1, ctx), "#008080");
		assert.equal(resolveColor(el2, ctx), "#80ffff");
	});

	it("switches to rgba() only once alpha is actually below opaque", () => {
		assert.equal(
			resolveColor(color("srgbClr", `val="FF0000"`, `<a:alpha val="50000"/>`), ctx),
			"rgba(255, 0, 0, 0.5)",
		);
		assert.equal(
			resolveColor(color("srgbClr", `val="FF0000"`, `<a:alpha val="100000"/>`), ctx),
			"#ff0000",
		);
	});

	it("ignores a transform it does not model instead of dropping the colour", () => {
		assert.equal(resolveColor(color("srgbClr", `val="FF0000"`, `<a:gamma/>`), ctx), "#ff0000");
	});
});

describe("resolveFillColor", () => {
	it("reaches through a fill wrapper to the colour inside", () => {
		assert.equal(
			resolveFillColor(el(`<a:solidFill xmlns:a="${A}"><a:srgbClr val="00FF00"/></a:solidFill>`), ctx),
			"#00ff00",
		);
	});

	it("returns null for a wrapper with nothing resolvable in it", () => {
		assert.equal(resolveFillColor(el(`<a:noFill xmlns:a="${A}"/>`), ctx), null);
		assert.equal(resolveFillColor(null, ctx), null);
	});
});

describe("parseTheme", () => {
	const theme = el(`<a:theme xmlns:a="${A}">
		<a:themeElements>
			<a:clrScheme name="Test">
				<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
				<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
				<a:accent1><a:srgbClr val="FF0000"/></a:accent1>
			</a:clrScheme>
			<a:fontScheme name="Test">
				<a:majorFont><a:latin typeface="Georgia"/></a:majorFont>
				<a:minorFont><a:latin typeface="Verdana"/></a:minorFont>
			</a:fontScheme>
			<a:fmtScheme>
				<a:fillStyleLst><a:solidFill/><a:gradFill/></a:fillStyleLst>
				<a:lnStyleLst><a:ln/></a:lnStyleLst>
				<a:bgFillStyleLst><a:solidFill/></a:bgFillStyleLst>
			</a:fmtScheme>
		</a:themeElements>
	</a:theme>`);

	it("reads the scheme slots the theme declares", () => {
		const parsed = parseTheme(theme, null);
		assert.equal(parsed.scheme.accent1, "#ff0000");
		assert.equal(parsed.scheme.dk1, "#000000");
	});

	it("keeps the default for a slot the theme leaves out", () => {
		assert.equal(parseTheme(theme, null).scheme.accent6, DEFAULT_THEME.scheme.accent6);
	});

	it("reads the major and minor latin typefaces", () => {
		const parsed = parseTheme(theme, null);
		assert.equal(parsed.majorFont, "Georgia");
		assert.equal(parsed.minorFont, "Verdana");
	});

	it("keeps the fmtScheme lists as elements, to resolve lazily", () => {
		const parsed = parseTheme(theme, null);
		assert.equal(parsed.fmt.fillStyles.length, 2);
		assert.equal(parsed.fmt.lineStyles.length, 1);
		assert.equal(parsed.fmt.bgFillStyles.length, 1);
	});

	it("takes the colour map off the master, so an inverted deck resolves inverted", () => {
		const clrMap = el(`<p:clrMap xmlns:p="${P}" bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2"/>`);
		const parsed = parseTheme(theme, clrMap);
		assert.equal(parsed.clrMap.bg1, "dk1");
		assert.equal(resolveColor(color("schemeClr", `val="bg1"`), { theme: parsed }), "#000000");
		assert.equal(resolveColor(color("schemeClr", `val="tx1"`), { theme: parsed }), "#ffffff");
	});

	it("falls back to the default theme when there is no theme part", () => {
		const parsed = parseTheme(null, null);
		assert.deepEqual(parsed.scheme, DEFAULT_THEME.scheme);
		assert.deepEqual(parsed.clrMap, DEFAULT_THEME.clrMap);
		assert.equal(parsed.majorFont, DEFAULT_THEME.majorFont);
		assert.deepEqual(parsed.fmt.fillStyles, []);
	});
});

describe("gradientCss", () => {
	function grad(inner: string): Element {
		return el(`<a:gradFill xmlns:a="${A}">${inner}</a:gradFill>`);
	}

	const stops = `<a:gsLst>
		<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
		<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
	</a:gsLst>`;

	it("turns an OOXML angle into the CSS angle for the same direction", () => {
		assert.equal(
			gradientCss(grad(`${stops}<a:lin ang="0"/>`), ctx),
			"linear-gradient(90deg, #ff0000 0%, #0000ff 100%)",
		);
	});

	it("defaults to a top-to-bottom gradient when no angle is given", () => {
		assert.equal(
			gradientCss(grad(stops), ctx),
			"linear-gradient(180deg, #ff0000 0%, #0000ff 100%)",
		);
	});

	it("sorts the stops by position, whatever order the file lists them in", () => {
		const reversed = grad(`<a:gsLst>
			<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
			<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
		</a:gsLst><a:lin ang="0"/>`);
		assert.equal(gradientCss(reversed, ctx), "linear-gradient(90deg, #ff0000 0%, #0000ff 100%)");
	});

	it("renders a path gradient as a radial one", () => {
		assert.equal(
			gradientCss(grad(`${stops}<a:path path="circle"/>`), ctx),
			"radial-gradient(circle at 50% 50%, #ff0000 0%, #0000ff 100%)",
		);
	});

	it("makes a single stop a flat fill rather than a one-sided gradient", () => {
		const one = grad(`<a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs></a:gsLst>`);
		assert.equal(gradientCss(one, ctx), "linear-gradient(#ff0000, #ff0000)");
	});

	it("returns null when no stop resolves, so the caller can leave the fill alone", () => {
		assert.equal(gradientCss(grad(`<a:gsLst/>`), ctx), null);
	});
});
