import { resolveFillColor } from "./color";
import type { PptxPackage } from "./package";
import type { ParseContext } from "./style";
import { colorContext } from "./style";
import { own } from "./lookup";
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

/**
 * One chart type within the plot area. Most charts have exactly one; a
 * combination chart — columns with a line over them — has one per type.
 */
export interface ChartGroup {
	kind: ChartKind;
	stacked: boolean;
	percentStacked: boolean;
	/** Plotted against the second value axis, which is drawn on the right. */
	secondary: boolean;
	/** Indices into `ChartModel.series`. */
	series: number[];
}

export interface ChartModel {
	/** The first group's type, which is what decides the chart's overall layout. */
	kind: ChartKind;
	title: string;
	categories: string[];
	/** Every series in the chart, across all its groups, in plot-area order. */
	series: ChartSeries[];
	/** The first group's stacking, as `kind` is the first group's type. */
	stacked: boolean;
	percentStacked: boolean;
	groups: ChartGroup[];
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
		groups: [],
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

	// Each chart type in the plot area is a group of its own; a combination
	// chart is just more than one of them.
	const plots: { el: Element; kind: ChartKind }[] = [];
	for (let n = plotArea.firstElementChild; n; n = n.nextElementSibling) {
		const mapped = own(KIND_BY_ELEMENT, n.localName);
		if (!mapped) continue;
		// A bar chart is drawn horizontally or vertically depending on c:barDir.
		const kind = mapped === "column" && attr(child(n, "barDir"), "val") === "bar" ? "bar" : mapped;
		plots.push({ el: n, kind });
	}
	if (plots.length === 0) return emptyChart();
	const plot = plots[0].el;

	const valueAxes = new Set(
		children(plotArea, "valAx")
			.map((ax) => attr(child(ax, "axId"), "val"))
			.filter((id): id is string => id !== null),
	);
	const primaryAxis = valueAxisOf(plot, valueAxes);

	const categories: string[] = [];
	const series: ChartSeries[] = [];
	const groups: ChartGroup[] = [];

	for (const { el, kind } of plots) {
		const grouping = attr(child(el, "grouping"), "val") ?? "clustered";
		const group: ChartGroup = {
			kind,
			stacked: grouping === "stacked" || grouping === "percentStacked",
			percentStacked: grouping === "percentStacked",
			secondary: primaryAxis !== null && valueAxisOf(el, valueAxes) !== primaryAxis,
			series: [],
		};
		for (const ser of children(el, "ser")) {
			const values = readNumbers(child(ser, "val") ?? child(ser, "yVal"));
			const cats = readStrings(child(ser, "cat") ?? child(ser, "xVal"));
			for (let i = 0; i < cats.length; i++) {
				if (cats[i] !== undefined && categories[i] === undefined) categories[i] = cats[i];
			}
			group.series.push(series.length);
			series.push({
				name: seriesName(ser, series.length),
				values,
				color: resolveFillColor(child(child(ser, "spPr"), "solidFill"), colorContext(ctx)),
				pointColors: readPointColors(ser, ctx),
			});
		}
		groups.push(group);
	}

	// Categories are optional; fall back to 1..n so a chart still has an axis.
	const length = Math.max(0, ...series.map((s) => s.values.length));
	for (let i = 0; i < length; i++) {
		if (categories[i] === undefined) categories[i] = String(i + 1);
	}

	const legendPos = attr(child(child(chart, "legend"), "legendPos"), "val");
	const hasLegend = child(chart, "legend") !== null;

	return {
		kind: groups[0].kind,
		title: readTitle(child(chart, "title")),
		categories: categories.slice(0, length),
		series,
		stacked: groups[0].stacked,
		percentStacked: groups[0].percentStacked,
		groups,
		legend: hasLegend ? normaliseLegend(legendPos) : "none",
		holeSize: (numAttr(child(plot, "holeSize"), "val") ?? 50) / 100,
		categoryAxisTitle: readTitle(child(findAxis(plotArea, "catAx", "dateAx"), "title")),
		valueAxisTitle: readTitle(child(findAxis(plotArea, "valAx"), "title")),
		empty: series.length === 0,
		palette: accentPalette(ctx),
	};
}

/**
 * The value axis a group plots against: the last of its c:axId that names a
 * c:valAx. A column or line group lists its category axis first; a scatter
 * group lists its x axis first, which is a value axis too, so taking the last
 * one finds the y axis in both.
 */
function valueAxisOf(plot: Element, valueAxes: Set<string>): string | null {
	const ids = children(plot, "axId")
		.map((el) => attr(el, "val"))
		.filter((id): id is string => id !== null && valueAxes.has(id));
	return ids[ids.length - 1] ?? null;
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
	const out: (number | null)[] = Array.from({ length: count ?? 0 }, () => null);
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
