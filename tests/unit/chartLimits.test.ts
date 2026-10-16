/**
 * How big a chart the plugin is willing to read.
 *
 * A chart's cached values say how many points a series has (`c:ptCount`) and
 * which point each value is (`c:pt@idx`), and the reader lays them out in an
 * array that long. Both are just numbers in the file, so a single point at
 * index two hundred million asks for an array of two hundred million — the
 * deck view stops responding while it is filled. Excel itself stopped capping
 * series length after 2007 (it was 32,000 points), and allows 255 series in
 * one chart; the limits here sit well above anything a slide can show, and a
 * chart past them is reported as too large instead of being drawn.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import { installDom } from "../../scripts/dom-shim";

installDom();

import { MAX_CHART_POINTS, MAX_SERIES_POINTS, parseChartPart, type ChartModel } from "../../src/pptx/chart";
import { DEFAULT_THEME } from "../../src/pptx/color";
import { renderChart } from "../../src/render/renderChart";
import type { ChartShape } from "../../src/pptx/types";
import { miniDeckParts, packageOf } from "./miniDeck";

const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

function chartOf(plot: string): ChartModel {
	const parts = miniDeckParts([{}]);
	parts["ppt/charts/chart1.xml"] =
		`<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart>` +
		"<c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>" +
		`<c:plotArea><c:barChart><c:barDir val="col"/>${plot}</c:barChart></c:plotArea></c:chart></c:chartSpace>`;
	const pkg = packageOf(parts);
	return parseChartPart(pkg, "ppt/charts/chart1.xml", { pkg, partPath: "ppt/slides/slide1.xml", theme: DEFAULT_THEME });
}

function ser(values: string, cats = "", extra = ""): string {
	return `<c:ser>${extra}${cats}<c:val><c:numLit>${values}</c:numLit></c:val></c:ser>`;
}

const pt = (idx: number | string, v: number | string = 1) => `<c:pt idx="${idx}"><c:v>${v}</c:v></c:pt>`;

describe("chart size limits", () => {
	it("are generous: 100,000 points a series, a million a chart", () => {
		assert.equal(MAX_SERIES_POINTS, 100_000);
		assert.equal(MAX_CHART_POINTS, 1_000_000);
	});

	it("still read an ordinary chart", () => {
		const chart = chartOf(ser(`<c:ptCount val="3"/>${pt(0, 1)}${pt(1, 2)}${pt(2, 3)}`));
		assert.equal(chart.tooLarge, false);
		assert.deepEqual(chart.series[0].values, [1, 2, 3]);
	});

	it("refuse a declared point count past the limit, without allocating it", () => {
		const started = Date.now();
		const chart = chartOf(ser(`<c:ptCount val="200000000"/>${pt(0)}`));
		assert.equal(chart.tooLarge, true);
		assert.equal(chart.series.length, 0);
		assert.ok(Date.now() - started < 2000, "took long enough to have filled the array");
	});

	it("refuse a single point placed past the limit", () => {
		assert.equal(chartOf(ser(`<c:ptCount val="1"/>${pt(200_000_000)}`)).tooLarge, true);
		assert.equal(chartOf(ser(`<c:ptCount val="1"/>${pt(MAX_SERIES_POINTS)}`)).tooLarge, true);
		assert.equal(chartOf(ser(`<c:ptCount val="1"/>${pt(MAX_SERIES_POINTS - 1)}`)).tooLarge, false);
	});

	it("refuse a category or a point colour placed past the limit", () => {
		const cats = `<c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="200000000"><c:v>Q1</c:v></c:pt></c:strLit></c:cat>`;
		assert.equal(chartOf(ser(`<c:ptCount val="1"/>${pt(0)}`, cats)).tooLarge, true);
		const dPt = '<c:dPt><c:idx val="200000000"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:dPt>';
		assert.equal(chartOf(ser(`<c:ptCount val="1"/>${pt(0)}`, "", dPt)).tooLarge, true);
	});

	it("refuse many series that are each within the limit but together are not", () => {
		const series = Array.from({ length: 11 }, () => ser(`<c:ptCount val="${MAX_SERIES_POINTS}"/>`)).join("");
		assert.equal(chartOf(series).tooLarge, true);
	});

	it("ignore point indices that are not whole, non-negative numbers", () => {
		const chart = chartOf(ser(`<c:ptCount val="2"/>${pt(-5, 9)}${pt("1.5", 9)}${pt("x", 9)}${pt(1, 4)}`));
		assert.equal(chart.tooLarge, false);
		assert.deepEqual(chart.series[0].values, [null, 4]);
	});

	it("draw a chart that is too large as a message, keeping its title", () => {
		const chart = chartOf(ser(`<c:ptCount val="200000000"/>`));
		const shape: ChartShape = {
			kind: "chart",
			id: "4",
			name: "Chart",
			frame: { x: 0, y: 0, w: 600, h: 400, rot: 0, flipH: false, flipV: false },
			hidden: false,
			source: null,
			sourcePart: "ppt/slides/slide1.xml",
			chart,
		};
		const texts = Array.from(renderChart(shape).querySelectorAll("text")).map((t) => t.textContent ?? "");
		assert.ok(texts.includes("Sales"), "the title went missing");
		assert.ok(texts.some((t) => /too many/i.test(t)), `no message among ${JSON.stringify(texts)}`);
	});
});
