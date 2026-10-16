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
	/** True when it has more points than are read (see MAX_SERIES_POINTS); it is then also empty. */
	tooLarge: boolean;
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
		tooLarge: false,
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
	const budget: Budget = { used: 0 };

	// A chart past the size limits is reported as such rather than read.
	try {
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
				const values = readNumbers(child(ser, "val") ?? child(ser, "yVal"), budget);
				const cats = readStrings(child(ser, "cat") ?? child(ser, "xVal"), budget);
				for (let i = 0; i < cats.length; i++) {
					if (cats[i] !== undefined && categories[i] === undefined) categories[i] = cats[i];
				}
				group.series.push(series.length);
				series.push({
					name: seriesName(ser, series.length),
					values,
					color: resolveFillColor(child(child(ser, "spPr"), "solidFill"), colorContext(ctx)),
					pointColors: readPointColors(ser, ctx, budget),
				});
			}
			groups.push(group);
		}
	} catch (error) {
		if (!(error instanceof ChartTooLarge)) throw error;
		return { ...emptyChart(), title: readTitle(child(chart, "title")), tooLarge: true };
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
		tooLarge: false,
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
 * The most points read for one series, and for the whole chart.
 *
 * A series is laid out in an array as long as its `c:ptCount`, or its highest
 * `c:pt@idx`, says — both just numbers in the file. Excel stopped capping series
 * length after 2007, where the limit was 32,000 points, and allows 255 series
 * in a chart; these sit well above anything a slide can show. A chart past
 * either is reported as too large rather than read, so a number in the file
 * cannot ask for an array big enough to stall the view.
 */
export const MAX_SERIES_POINTS = 100_000;
export const MAX_CHART_POINTS = 1_000_000;

class ChartTooLarge extends Error {}

/** Points laid out so far in the chart being read. */
interface Budget {
	used: number;
}

/** Make room for a series this long, or give up on the chart. */
function reserve(budget: Budget, length: number): void {
	if (length > MAX_SERIES_POINTS) throw new ChartTooLarge();
	budget.used += length;
	if (budget.used > MAX_CHART_POINTS) throw new ChartTooLarge();
}

/**
 * A point index read from the file. Absent counts as the first point, as it
 * always has; anything that is not a whole, non-negative number is not a point.
 */
function pointIndex(el: Element | null, name: string): number | null {
	const raw = attr(el, name);
	if (raw === null) return 0;
	const n = Number(raw);
	return Number.isInteger(n) && n >= 0 ? n : null;
}

/** The elements of a cache with their indices, and how long an array they need. */
function indexed(elements: Element[], indexOf: (el: Element) => number | null, declared = 0): {
	points: { el: Element; index: number }[];
	length: number;
} {
	const points: { el: Element; index: number }[] = [];
	let length = declared;
	for (const el of elements) {
		const index = indexOf(el);
		if (index === null) continue;
		points.push({ el, index });
		if (index + 1 > length) length = index + 1;
	}
	return { points, length };
}

/**
 * Read a cached numeric column. Points carry their own index, and gaps are real
 * — a missing point is a hole in the line, not a zero.
 */
function readNumbers(container: Element | null, budget: Budget): (number | null)[] {
	const cache = child(child(container, "numRef"), "numCache") ?? child(container, "numLit");
	if (!cache) return [];
	const count = numAttr(child(cache, "ptCount"), "val");
	const declared = count !== null && Number.isInteger(count) && count > 0 ? count : 0;
	if (declared > MAX_SERIES_POINTS) throw new ChartTooLarge();
	const { points, length } = indexed(children(cache, "pt"), (pt) => pointIndex(pt, "idx"), declared);
	reserve(budget, length);
	const out: (number | null)[] = Array.from({ length }, () => null);
	for (const { el: pt, index } of points) {
		const raw = child(pt, "v")?.textContent?.trim();
		const value = raw === undefined || raw === "" ? null : Number(raw);
		out[index] = value !== null && Number.isFinite(value) ? value : null;
	}
	return out;
}

function readStrings(container: Element | null, budget: Budget): string[] {
	const cache =
		child(child(container, "strRef"), "strCache") ??
		child(child(container, "numRef"), "numCache") ??
		child(container, "strLit") ??
		child(container, "numLit") ??
		// Multi-level categories collapse to their innermost level.
		child(child(child(container, "multiLvlStrRef"), "multiLvlStrCache"), "lvl");
	if (!cache) return [];
	const { points, length } = indexed(children(cache, "pt"), (pt) => pointIndex(pt, "idx"));
	reserve(budget, length);
	const out: string[] = Array.from({ length }, () => "");
	for (const { el: pt, index } of points) out[index] = child(pt, "v")?.textContent ?? "";
	return out;
}

/** Explicit colours on individual data points, which pie charts rely on. */
function readPointColors(ser: Element, ctx: ParseContext, budget: Budget): (string | null)[] {
	const { points, length } = indexed(children(ser, "dPt"), (dPt) => pointIndex(child(dPt, "idx"), "val"));
	reserve(budget, length);
	const out: (string | null)[] = Array.from({ length }, () => null);
	for (const { el: dPt, index } of points) {
		out[index] = resolveFillColor(child(child(dPt, "spPr"), "solidFill"), colorContext(ctx));
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
