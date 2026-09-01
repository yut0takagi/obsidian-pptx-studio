/**
 * Frames and preset shapes.
 *
 * `parseFrame` is where every EMU in a deck becomes a pixel, so the conversion
 * and the two odd units around it — rotation in sixtieths of a degree, flips as
 * OOXML booleans — are pinned here rather than left to a rendering diff.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import { installDomParser } from "./dom";

installDomParser();

import {
	EMPTY_FRAME,
	geometryCss,
	geometryName,
	isLineGeometry,
	parseChildFrame,
	parseFrame,
} from "../../src/pptx/geometry";
import { EMU_PER_PX } from "../../src/pptx/types";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

function el(xml: string): Element {
	return new DOMParser().parseFromString(xml, "application/xml").documentElement;
}

function xfrm(inner: string, attrs = ""): Element {
	return el(`<a:xfrm xmlns:a="${A}" ${attrs}>${inner}</a:xfrm>`);
}

describe("parseFrame", () => {
	it("converts EMU to pixels", () => {
		const frame = parseFrame(
			xfrm(`<a:off x="${EMU_PER_PX * 100}" y="${EMU_PER_PX * 50}"/>
			      <a:ext cx="${EMU_PER_PX * 300}" cy="${EMU_PER_PX * 200}"/>`),
		);
		assert.deepEqual(frame, { x: 100, y: 50, w: 300, h: 200, rot: 0, flipH: false, flipV: false });
	});

	it("reads rotation in sixtieths of a degree", () => {
		assert.equal(parseFrame(xfrm("", `rot="5400000"`))?.rot, 90);
		assert.equal(parseFrame(xfrm("", `rot="-2700000"`))?.rot, -45);
	});

	it("reads the flips", () => {
		const frame = parseFrame(xfrm("", `flipH="1" flipV="0"`));
		assert.equal(frame?.flipH, true);
		assert.equal(frame?.flipV, false);
	});

	it("treats a missing off or ext as zero rather than NaN", () => {
		assert.deepEqual(parseFrame(xfrm("")), EMPTY_FRAME);
	});

	it("returns null when there is no xfrm at all, so inheritance can kick in", () => {
		assert.equal(parseFrame(null), null);
	});
});

describe("parseChildFrame", () => {
	it("reads the child coordinate space a group declares", () => {
		const frame = parseChildFrame(
			xfrm(`<a:chOff x="${EMU_PER_PX * 10}" y="0"/>
			      <a:chExt cx="${EMU_PER_PX * 400}" cy="${EMU_PER_PX * 300}"/>`),
		);
		assert.deepEqual(frame, { x: 10, y: 0, w: 400, h: 300 });
	});

	it("returns null when the xfrm declares no child space", () => {
		assert.equal(parseChildFrame(xfrm(`<a:off x="0" y="0"/>`)), null);
		assert.equal(parseChildFrame(null), null);
	});
});

describe("geometryName", () => {
	it("reads the preset off prstGeom", () => {
		const spPr = el(`<a:spPr xmlns:a="${A}"><a:prstGeom prst="roundRect"/></a:spPr>`);
		assert.equal(geometryName(spPr), "roundRect");
	});

	it("names a custom geometry so the renderer can tell it apart from a plain box", () => {
		const spPr = el(`<a:spPr xmlns:a="${A}"><a:custGeom/></a:spPr>`);
		assert.equal(geometryName(spPr), "custGeom");
	});

	it("falls back to a rectangle when the shape declares nothing", () => {
		assert.equal(geometryName(el(`<a:spPr xmlns:a="${A}"/>`)), "rect");
		assert.equal(geometryName(null), "rect");
	});
});

describe("geometryCss", () => {
	it("leaves a rectangle unstyled, so it keeps a real CSS border", () => {
		assert.deepEqual(geometryCss("rect", 100, 50), { clipped: false });
	});

	it("scales a rounded corner off the shorter side", () => {
		assert.deepEqual(geometryCss("roundRect", 300, 100), { borderRadius: "16px", clipped: false });
	});

	it("rounds a terminator into a full stadium", () => {
		assert.deepEqual(geometryCss("flowChartTerminator", 300, 80), {
			borderRadius: "40px",
			clipped: false,
		});
	});

	it("marks clip-path shapes as clipped, since a CSS border would show through", () => {
		const triangle = geometryCss("triangle", 100, 100);
		assert.equal(triangle.clipped, true);
		assert.equal(triangle.clipPath, "polygon(50% 0%, 100% 100%, 0% 100%)");
		assert.equal(triangle.borderRadius, undefined);
	});

	it("keeps an ellipse a border rather than a clip", () => {
		assert.deepEqual(geometryCss("ellipse", 100, 100), { borderRadius: "50%", clipped: false });
	});

	it("falls back to a rectangle for a preset it does not model", () => {
		assert.deepEqual(geometryCss("noSmoking", 100, 100), { clipped: false });
	});
});

describe("isLineGeometry", () => {
	it("recognises lines and every flavour of connector", () => {
		for (const name of [
			"line",
			"straightConnector1",
			"bentConnector2",
			"bentConnector3",
			"curvedConnector4",
		]) {
			assert.equal(isLineGeometry(name), true, name);
		}
	});

	it("leaves boxes alone", () => {
		for (const name of ["rect", "roundRect", "ellipse", "custGeom"]) {
			assert.equal(isLineGeometry(name), false, name);
		}
	});
});
