/**
 * Combination charts: columns and a line in one plot area.
 *
 * A chart's plot area holds one element per chart type, and a combination
 * chart is simply more than one of them — typically columns for an amount and
 * a line for a rate, the line on a second value axis because the two scales
 * have nothing in common. Reading only the first group drops the line
 * altogether; drawing the line against the columns' axis flattens it onto the
 * baseline. Both are checked here, along with the line sitting over the
 * middle of each group of columns rather than on the category edges.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import { installDom } from "../../scripts/dom-shim";

installDom();

import { parseChartPart, type ChartModel } from "../../src/pptx/chart";
import { DEFAULT_THEME } from "../../src/pptx/color";
import { renderChart } from "../../src/render/renderChart";
import type { ChartShape } from "../../src/pptx/types";
import { miniDeckParts, packageOf } from "./miniDeck";

const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

function ser(idx: number, name: string, colour: string, values: number[]): string {
	const cats = ["Q1", "Q2", "Q3", "Q4"];
	return (
		`<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/>` +
		`<c:tx><c:v>${name}</c:v></c:tx>` +
		`<c:spPr><a:solidFill><a:srgbClr val="${colour}"/></a:solidFill></c:spPr>` +
		`<c:cat><c:strLit><c:ptCount val="4"/>${cats.map((c, i) => `<c:pt idx="${i}"><c:v>${c}</c:v></c:pt>`).join("")}</c:strLit></c:cat>` +
		`<c:val><c:numLit><c:ptCount val="4"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("")}</c:numLit></c:val>` +
		"</c:ser>"
	);
}

/** Revenue as columns on the primary axis, margin as a line on a secondary one. */
const COMBO =
	`<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:plotArea>` +
	'<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>' +
	ser(0, "Revenue", "111111", [100, 300, 200, 500]) +
	ser(1, "Cost", "222222", [80, 200, 150, 300]) +
	'<c:axId val="1"/><c:axId val="2"/></c:barChart>' +
	'<c:lineChart><c:grouping val="standard"/>' +
	ser(2, "Margin", "333333", [0.1, 0.2, 0.3, 0.5]) +
	'<c:axId val="3"/><c:axId val="4"/></c:lineChart>' +
	'<c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
	'<c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>' +
	'<c:catAx><c:axId val="3"/><c:axPos val="b"/><c:delete val="1"/><c:crossAx val="4"/></c:catAx>' +
	'<c:valAx><c:axId val="4"/><c:axPos val="r"/><c:crossAx val="3"/></c:valAx>' +
	"</c:plotArea><c:legend><c:legendPos val=\"b\"/></c:legend></c:chart></c:chartSpace>";

/** Columns and a line sharing one value axis. */
const SHARED =
	`<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:plotArea>` +
	'<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/>' +
	ser(0, "A", "111111", [1, 2, 3, 4]) +
	'<c:axId val="1"/><c:axId val="2"/></c:barChart>' +
	'<c:lineChart><c:grouping val="standard"/>' +
	ser(1, "B", "333333", [4, 3, 2, 1]) +
	'<c:axId val="1"/><c:axId val="2"/></c:lineChart>' +
	'<c:catAx><c:axId val="1"/></c:catAx><c:valAx><c:axId val="2"/></c:valAx>' +
	"</c:plotArea></c:chart></c:chartSpace>";

function chartOf(xml: string): ChartModel {
	const parts = miniDeckParts([{}]);
	parts["ppt/charts/chart1.xml"] = xml;
	const pkg = packageOf(parts);
	return parseChartPart(pkg, "ppt/charts/chart1.xml", {
		pkg,
		partPath: "ppt/slides/slide1.xml",
		theme: DEFAULT_THEME,
	});
}

function draw(chart: ChartModel): SVGElement {
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
	return renderChart(shape);
}

const num = (el: Element, name: string): number => Number(el.getAttribute(name));

describe("parsing a combination chart", () => {
	it("reads every group, not just the first", () => {
		const chart = chartOf(COMBO);
		assert.deepEqual(
			chart.series.map((s) => s.name),
			["Revenue", "Cost", "Margin"],
		);
		assert.equal(chart.groups.length, 2);
		assert.equal(chart.groups[0].kind, "column");
		assert.deepEqual(chart.groups[0].series, [0, 1]);
		assert.equal(chart.groups[1].kind, "line");
		assert.deepEqual(chart.groups[1].series, [2]);
		// The chart as a whole still answers to its first group.
		assert.equal(chart.kind, "column");
	});

	it("knows which group is on the second value axis", () => {
		const chart = chartOf(COMBO);
		assert.equal(chart.groups[0].secondary, false);
		assert.equal(chart.groups[1].secondary, true);
		const shared = chartOf(SHARED);
		assert.deepEqual(
			shared.groups.map((g) => g.secondary),
			[false, false],
		);
	});

	it("keeps each group's own stacking", () => {
		const shared = chartOf(SHARED);
		assert.equal(shared.groups[0].stacked, true);
		assert.equal(shared.groups[1].stacked, false);
	});
});

describe("drawing a combination chart", () => {
	it("draws the columns and the line", () => {
		const svg = draw(chartOf(COMBO));
		assert.equal(svg.querySelectorAll('rect[fill="#111111"]').length >= 4, true, "revenue columns");
		assert.equal(svg.querySelectorAll('rect[fill="#222222"]').length >= 4, true, "cost columns");
		assert.equal(svg.querySelectorAll('polyline[stroke="#333333"]').length, 1, "the margin line");
	});

	it("scales the line against its own axis rather than flattening it", () => {
		const svg = draw(chartOf(COMBO));
		const columnTop = Math.min(
			...Array.from(svg.querySelectorAll('rect[fill="#111111"]')).map((r) => num(r, "y")),
		);
		const lineTop = Math.min(
			...Array.from(svg.querySelectorAll('circle[fill="#333333"]')).map((c) => num(c, "cy")),
		);
		// Both peak at the top of their own axis: 500 of 500, and 0.5 of 0.5.
		assert.ok(Math.abs(columnTop - lineTop) < 1, `line peaks at ${lineTop}, columns at ${columnTop}`);
	});

	it("labels the second axis on the right", () => {
		const svg = draw(chartOf(COMBO));
		const right = Array.from(svg.querySelectorAll("text")).filter(
			(t) => t.getAttribute("text-anchor") === "start" && num(t, "x") > 400,
		);
		assert.ok(right.some((t) => t.textContent === "0.5"), "no 0.5 on the right-hand axis");
	});

	it("puts each point of the line over the middle of its group of columns", () => {
		const svg = draw(chartOf(COMBO));
		const firstPoint = svg.querySelector('circle[fill="#333333"]')!;
		const revenue = svg.querySelector('rect[fill="#111111"]')!;
		const cost = svg.querySelector('rect[fill="#222222"]')!;
		// Two clustered columns: the middle of the pair is where the second begins.
		assert.ok(Math.abs(num(firstPoint, "cx") - num(cost, "x")) < 0.01);
		assert.ok(num(revenue, "x") < num(firstPoint, "cx"));
	});

	it("lists every series in the legend", () => {
		const svg = draw(chartOf(COMBO));
		const labels = Array.from(svg.querySelectorAll("text")).map((t) => t.textContent);
		for (const name of ["Revenue", "Cost", "Margin"]) assert.ok(labels.includes(name), name);
	});
});
