import { resolveFillColor } from "./color";
import type { PptxPackage } from "./package";
import type { ParseContext } from "./style";
import { colorContext } from "./style";
import { attr, child, children, numAttr } from "./xml";

/**
 * Charts, read from the cached values every chart part carries.
 *
 * A chart's live data lives in an embedded workbook, but PowerPoint also writes
 * the last computed values into `c:numCache` / `c:strCache` so a consumer can
 * draw the chart without opening a spreadsheet. That cache is what this reads —
 * it is exactly what the deck was showing when it was saved.
 */

export type ChartKind =
	| "column"
	| "bar"
	| "line"
	| "area"
	| "pie"
	| "doughnut"
	| "scatter"
	| "radar"
	| "unknown";

export interface ChartSeries {
	name: string;
	/** Sparse points are filled with null so category indices stay aligned. */
	values: (number | null)[];
	/** Explicit series colour, when the chart declares one. */
	color: string | null;
	/** Per-point colours, used by pie and doughnut charts. */
	pointColors: (string | null)[];
}

export interface ChartModel {
	kind: ChartKind;
	title: string;
	categories: string[];
	series: ChartSeries[];
	stacked: boolean;
	percentStacked: boolean;
	legend: "none" | "b" | "t" | "l" | "r";
	/** Doughnut hole as a fraction of the radius. */
	holeSize: number;
	/** Category and value axis titles, when the chart names them. */
	categoryAxisTitle: string;
	valueAxisTitle: string;
	/** True when the chart part could not be understood at all. */
	empty: boolean;
	/**
	 * Default series colours, taken from the deck's theme accents. PowerPoint
	 * colours series from the accent ramp unless the chart overrides it, so a
	 * chart that declares nothing still comes out in the deck's palette.
	 */
	palette: string[];
}

const KIND_BY_ELEMENT: Record<string, ChartKind> = {
	barChart: "column",
	bar3DChart: "column",
	lineChart: "line",
	line3DChart: "line",
	areaChart: "area",
	area3DChart: "area",
	pieChart: "pie",
	pie3DChart: "pie",
	ofPieChart: "pie",
	doughnutChart: "doughnut",
	scatterChart: "scatter",
	bubbleChart: "scatter",
	radarChart: "radar",
};

export function emptyChart(): ChartModel {
	return {
		kind: "unknown",
		title: "",
		categories: [],
		series: [],
		stacked: false,
		percentStacked: false,
		legend: "b",
		holeSize: 0.5,
		categoryAxisTitle: "",
		valueAxisTitle: "",
		empty: true,
		palette: [],
	};
}

/** Parse a chart part into something the renderer can draw. */
export function parseChartPart(
	pkg: PptxPackage,
	chartPath: string,
	ctx: ParseContext,
): ChartModel {
	const root = pkg.xml(chartPath)?.documentElement ?? null;
	const chart = child(root, "chart");
	const plotArea = child(chart, "plotArea");
	if (!plotArea) return emptyChart();

	let kind: ChartKind = "unknown";
	let plot: Element | null = null;
	for (let n = plotArea.firstElementChild; n; n = n.nextElementSibling) {
		const mapped = KIND_BY_ELEMENT[n.localName];
		if (mapped) {
			kind = mapped;
			plot = n;
			break;
		}
	}
	if (!plot) return emptyChart();

	// A bar chart is drawn horizontally or vertically depending on c:barDir.
	if (kind === "column" && attr(child(plot, "barDir"), "val") === "bar") kind = "bar";

	const grouping = attr(child(plot, "grouping"), "val") ?? "clustered";
	const stacked = grouping === "stacked" || grouping === "percentStacked";
	const percentStacked = grouping === "percentStacked";

	const seriesEls = children(plot, "ser");
	const categories: string[] = [];
	const series: ChartSeries[] = [];

	for (const ser of seriesEls) {
		const values = readNumbers(child(ser, "val") ?? child(ser, "yVal"));
		const cats = readStrings(child(ser, "cat") ?? child(ser, "xVal"));
		for (let i = 0; i < cats.length; i++) {
			if (cats[i] !== undefined && categories[i] === undefined) categories[i] = cats[i];
		}
		series.push({
			name: seriesName(ser, series.length),
			values,
			color: resolveFillColor(child(child(ser, "spPr"), "solidFill"), colorContext(ctx)),
			pointColors: readPointColors(ser, ctx),
		});
	}

	// Categories are optional; fall back to 1..n so a chart still has an axis.
	const length = Math.max(0, ...series.map((s) => s.values.length));
	for (let i = 0; i < length; i++) {
		if (categories[i] === undefined) categories[i] = String(i + 1);
	}

	const legendPos = attr(child(child(chart, "legend"), "legendPos"), "val");
	const hasLegend = child(chart, "legend") !== null;

	return {
		kind,
		title: readTitle(child(chart, "title")),
		categories: categories.slice(0, length),
		series,
		stacked,
		percentStacked,
		legend: hasLegend ? normaliseLegend(legendPos) : "none",
		holeSize: (numAttr(child(plot, "holeSize"), "val") ?? 50) / 100,
		categoryAxisTitle: readTitle(child(findAxis(plotArea, "catAx", "dateAx"), "title")),
		valueAxisTitle: readTitle(child(findAxis(plotArea, "valAx"), "title")),
		empty: series.length === 0,
		palette: accentPalette(ctx),
	};
}

function accentPalette(ctx: ParseContext): string[] {
	const scheme = ctx.theme.scheme;
	return ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]
		.map((slot) => scheme[slot])
		.filter((color): color is string => Boolean(color));
}

function findAxis(plotArea: Element, ...names: string[]): Element | null {
	for (const name of names) {
		const found = child(plotArea, name);
		if (found) return found;
	}
	return null;
}

function normaliseLegend(value: string | null): ChartModel["legend"] {
	switch (value) {
		case "t":
		case "l":
		case "r":
			return value;
		case "tr":
			return "r";
		default:
			return "b";
	}
}

function seriesName(ser: Element, index: number): string {
	const cached = child(child(child(ser, "tx"), "strRef"), "strCache");
	const first = children(cached, "pt")[0];
	const text = child(first, "v")?.textContent?.trim();
	if (text) return text;
	const literal = child(child(ser, "tx"), "v")?.textContent?.trim();
	return literal || `Series ${index + 1}`;
}

/**
 * Read a cached numeric column. Points carry their own index, and gaps are real
 * — a missing point is a hole in the line, not a zero.
 */
function readNumbers(container: Element | null): (number | null)[] {
	const cache = child(child(container, "numRef"), "numCache") ?? child(container, "numLit");
	if (!cache) return [];
	const count = numAttr(child(cache, "ptCount"), "val");
	const out: (number | null)[] = new Array(count ?? 0).fill(null);
	for (const pt of children(cache, "pt")) {
		const index = numAttr(pt, "idx") ?? 0;
		const raw = child(pt, "v")?.textContent?.trim();
		const value = raw === undefined || raw === "" ? null : Number(raw);
		while (out.length <= index) out.push(null);
		out[index] = value !== null && Number.isFinite(value) ? value : null;
	}
	return out;
}

function readStrings(container: Element | null): string[] {
	const cache =
		child(child(container, "strRef"), "strCache") ??
		child(child(container, "numRef"), "numCache") ??
		child(container, "strLit") ??
		child(container, "numLit") ??
		// Multi-level categories collapse to their innermost level.
		child(child(child(container, "multiLvlStrRef"), "multiLvlStrCache"), "lvl");
	if (!cache) return [];
	const out: string[] = [];
	for (const pt of children(cache, "pt")) {
		const index = numAttr(pt, "idx") ?? 0;
		const text = child(pt, "v")?.textContent ?? "";
		while (out.length <= index) out.push("");
		out[index] = text;
	}
	return out;
}

/** Explicit colours on individual data points, which pie charts rely on. */
function readPointColors(ser: Element, ctx: ParseContext): (string | null)[] {
	const out: (string | null)[] = [];
	for (const dPt of children(ser, "dPt")) {
		const index = numAttr(child(dPt, "idx"), "val") ?? 0;
		const color = resolveFillColor(child(child(dPt, "spPr"), "solidFill"), colorContext(ctx));
		while (out.length <= index) out.push(null);
		out[index] = color;
	}
	return out;
}

/** The visible text of a c:title, from either rich text or a cached reference. */
function readTitle(title: Element | null): string {
	if (!title) return "";
	const rich = child(child(title, "tx"), "rich");
	if (rich) {
		return children(rich, "p")
			.map((p) => p.textContent ?? "")
			.join(" ")
			.trim();
	}
	const cached = child(child(child(title, "tx"), "strRef"), "strCache");
	if (cached) return (cached.textContent ?? "").trim();
	return "";
}
