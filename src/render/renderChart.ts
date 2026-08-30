import type { ChartModel, ChartSeries } from "../pptx/chart";
import type { ChartShape } from "../pptx/types";

const SVG = "http://www.w3.org/2000/svg";

const AXIS_COLOUR = "rgba(0,0,0,0.35)";
const GRID_COLOUR = "rgba(0,0,0,0.10)";
const TEXT_COLOUR = "rgba(0,0,0,0.72)";
const FAINT_TEXT = "rgba(0,0,0,0.55)";

/**
 * Draw a chart as inline SVG.
 *
 * Everything is laid out in the shape's own pixel box rather than a fixed
 * viewBox, so a chart placed at 200px and one at 900px both get axis labels and
 * gridlines at a readable size instead of one being a scaled-up blur of the
 * other.
 */
export function renderChart(shape: ChartShape): SVGElement {
	const { chart } = shape;
	const width = Math.max(shape.frame.w, 40);
	const height = Math.max(shape.frame.h, 40);

	const svg = document.createElementNS(SVG, "svg");
	svg.setAttribute("width", "100%");
	svg.setAttribute("height", "100%");
	svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
	svg.style.overflow = "visible";

	const font = clamp(Math.min(width, height) / 24, 8, 15);
	const layout = layoutRegions(chart, width, height, font);

	if (chart.title) {
		text(svg, width / 2, layout.titleBaseline, chart.title, {
			size: font * 1.25,
			weight: "600",
			anchor: "middle",
			fill: TEXT_COLOUR,
		});
	}

	if (chart.empty || chart.series.length === 0) {
		text(svg, width / 2, height / 2, chart.title || "Chart", {
			size: font,
			anchor: "middle",
			fill: FAINT_TEXT,
		});
		return svg;
	}

	switch (chart.kind) {
		case "pie":
		case "doughnut":
			drawPie(svg, chart, layout, font);
			break;
		case "bar":
			drawBars(svg, chart, layout, font, true);
			break;
		case "line":
		case "radar":
			drawLines(svg, chart, layout, font, false);
			break;
		case "area":
			drawLines(svg, chart, layout, font, true);
			break;
		case "scatter":
			drawScatter(svg, chart, layout, font);
			break;
		default:
			drawBars(svg, chart, layout, font, false);
			break;
	}

	drawLegend(svg, chart, layout, font);
	return svg;
}

// ---------------------------------------------------------------- layout

interface Layout {
	plot: { x: number; y: number; w: number; h: number };
	legend: { x: number; y: number; w: number; h: number } | null;
	titleBaseline: number;
}

function layoutRegions(chart: ChartModel, width: number, height: number, font: number): Layout {
	const pad = font * 0.7;
	let top = pad;
	let bottom = height - pad;
	let left = pad;
	let right = width - pad;

	const titleBaseline = top + font * 1.25;
	if (chart.title) top += font * 2.1;

	let legend: Layout["legend"] = null;
	if (chart.legend !== "none" && chart.series.length > 0) {
		const band = font * 2;
		switch (chart.legend) {
			case "r": {
				const w = Math.min(width * 0.3, longestLabel(chart) * font * 0.62 + font * 2);
				legend = { x: right - w, y: top, w, h: bottom - top };
				right -= w;
				break;
			}
			case "l": {
				const w = Math.min(width * 0.3, longestLabel(chart) * font * 0.62 + font * 2);
				legend = { x: left, y: top, w, h: bottom - top };
				left += w;
				break;
			}
			case "t":
				legend = { x: left, y: top, w: right - left, h: band };
				top += band;
				break;
			default:
				legend = { x: left, y: bottom - band, w: right - left, h: band };
				bottom -= band;
				break;
		}
	}

	// Category charts need gutters for their axis labels; pies do not.
	if (chart.kind !== "pie" && chart.kind !== "doughnut") {
		const ticks = valueTicks(chart);
		const widest = Math.max(...ticks.map((v) => formatTick(v).length), 3);
		if (chart.kind === "bar") {
			left += Math.max(...chart.categories.map((c) => c.length), 3) * font * 0.55 + font * 0.5;
			bottom -= font * 1.7;
		} else {
			left += widest * font * 0.6 + font * 0.5;
			bottom -= font * 1.7;
		}
	}

	return {
		plot: { x: left, y: top, w: Math.max(right - left, 10), h: Math.max(bottom - top, 10) },
		legend,
		titleBaseline,
	};
}

function longestLabel(chart: ChartModel): number {
	const source =
		chart.kind === "pie" || chart.kind === "doughnut" ? chart.categories : chart.series.map((s) => s.name);
	return Math.max(...source.map((s) => s.length), 4);
}

// ----------------------------------------------------------------- scale

/** Value range across the chart, taking stacking into account. */
function valueRange(chart: ChartModel): { min: number; max: number } {
	if (chart.percentStacked) return { min: 0, max: 1 };
	let min = 0;
	let max = 0;
	if (chart.stacked) {
		const length = Math.max(...chart.series.map((s) => s.values.length), 0);
		for (let i = 0; i < length; i++) {
			let positive = 0;
			let negative = 0;
			for (const series of chart.series) {
				const value = series.values[i] ?? 0;
				if (value >= 0) positive += value;
				else negative += value;
			}
			max = Math.max(max, positive);
			min = Math.min(min, negative);
		}
	} else {
		for (const series of chart.series) {
			for (const value of series.values) {
				if (value === null) continue;
				max = Math.max(max, value);
				min = Math.min(min, value);
			}
		}
	}
	if (min === 0 && max === 0) return { min: 0, max: 1 };
	return { min, max };
}

/** Round tick values, chosen so the axis reads in human numbers. */
function valueTicks(chart: ChartModel): number[] {
	const { min, max } = valueRange(chart);
	const span = max - min || 1;
	const rough = span / 4;
	const magnitude = 10 ** Math.floor(Math.log10(rough));
	const normalised = rough / magnitude;
	const step = (normalised >= 5 ? 10 : normalised >= 2.5 ? 5 : normalised >= 1.5 ? 2 : 1) * magnitude;

	const start = Math.floor(min / step) * step;
	const end = Math.ceil(max / step) * step;
	const ticks: number[] = [];
	for (let v = start; v <= end + step / 1000; v += step) {
		ticks.push(Math.abs(v) < step / 1000 ? 0 : v);
	}
	return ticks;
}

function formatTick(value: number): string {
	const abs = Math.abs(value);
	if (abs >= 1_000_000) return `${trim(value / 1_000_000)}M`;
	if (abs >= 1000) return `${trim(value / 1000)}k`;
	if (abs > 0 && abs < 1) return trim(value, 2);
	return trim(value);
}

function trim(value: number, digits = 1): string {
	const rounded = Number(value.toFixed(digits));
	return String(rounded);
}

// ------------------------------------------------------------ cartesian

interface Axes {
	ticks: number[];
	toValue: (v: number) => number;
	zero: number;
}

function drawCartesianFrame(
	svg: SVGElement,
	chart: ChartModel,
	layout: Layout,
	font: number,
	horizontal: boolean,
): Axes {
	const { plot } = layout;
	const ticks = valueTicks(chart);
	const lo = ticks[0];
	const hi = ticks[ticks.length - 1];
	const span = hi - lo || 1;

	const toValue = (v: number): number =>
		horizontal ? plot.x + ((v - lo) / span) * plot.w : plot.y + plot.h - ((v - lo) / span) * plot.h;

	for (const tick of ticks) {
		const at = toValue(tick);
		if (horizontal) {
			line(svg, at, plot.y, at, plot.y + plot.h, tick === 0 ? AXIS_COLOUR : GRID_COLOUR);
			text(svg, at, plot.y + plot.h + font * 1.25, formatValue(chart, tick), {
				size: font * 0.85,
				anchor: "middle",
				fill: FAINT_TEXT,
			});
		} else {
			line(svg, plot.x, at, plot.x + plot.w, at, tick === 0 ? AXIS_COLOUR : GRID_COLOUR);
			text(svg, plot.x - font * 0.4, at + font * 0.3, formatValue(chart, tick), {
				size: font * 0.85,
				anchor: "end",
				fill: FAINT_TEXT,
			});
		}
	}

	return { ticks, toValue, zero: toValue(Math.min(Math.max(0, lo), hi)) };
}

function formatValue(chart: ChartModel, value: number): string {
	return chart.percentStacked ? `${Math.round(value * 100)}%` : formatTick(value);
}

function seriesColour(chart: ChartModel, series: ChartSeries, index: number): string {
	if (series.color) return series.color;
	const palette = chart.palette.length > 0 ? chart.palette : ["#4472c4", "#ed7d31", "#a5a5a5"];
	return palette[index % palette.length];
}

// -------------------------------------------------------------- columns

function drawBars(
	svg: SVGElement,
	chart: ChartModel,
	layout: Layout,
	font: number,
	horizontal: boolean,
): void {
	const { plot } = layout;
	const axes = drawCartesianFrame(svg, chart, layout, font, horizontal);
	const count = Math.max(chart.categories.length, 1);
	const band = (horizontal ? plot.h : plot.w) / count;
	const groupWidth = band * 0.72;
	const barWidth = chart.stacked ? groupWidth : groupWidth / Math.max(chart.series.length, 1);

	for (let c = 0; c < count; c++) {
		const bandStart = (horizontal ? plot.y : plot.x) + c * band + (band - groupWidth) / 2;
		let positiveTop = 0;
		let negativeTop = 0;

		chart.series.forEach((series, s) => {
			let value = series.values[c];
			if (value === null || value === undefined) return;
			if (chart.percentStacked) {
				const total = chart.series.reduce((sum, other) => sum + Math.abs(other.values[c] ?? 0), 0);
				value = total === 0 ? 0 : value / total;
			}

			let from: number;
			let to: number;
			if (chart.stacked) {
				const base = value >= 0 ? positiveTop : negativeTop;
				from = axes.toValue(base);
				to = axes.toValue(base + value);
				if (value >= 0) positiveTop += value;
				else negativeTop += value;
			} else {
				from = axes.zero;
				to = axes.toValue(value);
			}

			const offset = chart.stacked ? 0 : s * barWidth;
			const rect = document.createElementNS(SVG, "rect");
			if (horizontal) {
				rect.setAttribute("x", String(Math.min(from, to)));
				rect.setAttribute("y", String(bandStart + offset));
				rect.setAttribute("width", String(Math.max(Math.abs(to - from), 0.5)));
				rect.setAttribute("height", String(Math.max(barWidth * 0.86, 1)));
			} else {
				rect.setAttribute("x", String(bandStart + offset));
				rect.setAttribute("y", String(Math.min(from, to)));
				rect.setAttribute("width", String(Math.max(barWidth * 0.86, 1)));
				rect.setAttribute("height", String(Math.max(Math.abs(to - from), 0.5)));
			}
			rect.setAttribute("fill", seriesColour(chart, series, s));
			svg.appendChild(rect);
		});
	}

	drawCategoryLabels(svg, chart, layout, font, horizontal, band, (i) => band * (i + 0.5));
}

/**
 * Category labels along the flat axis.
 *
 * Bars sit inside a band and label its centre; line and area points sit *on* the
 * division, so the caller supplies where its category i actually is rather than
 * this assuming one convention and being half a band out for the other.
 */
function drawCategoryLabels(
	svg: SVGElement,
	chart: ChartModel,
	layout: Layout,
	font: number,
	horizontal: boolean,
	spacing: number,
	offsetOf: (index: number) => number,
): void {
	const { plot } = layout;
	// With more labels than room, show every nth so they stay readable.
	const minSpacing = horizontal ? font * 1.4 : font * 3;
	const stride = Math.max(1, Math.ceil(minSpacing / Math.max(spacing, 1)));

	chart.categories.forEach((label, i) => {
		if (i % stride !== 0) return;
		if (horizontal) {
			text(svg, plot.x - font * 0.4, plot.y + offsetOf(i) + font * 0.3, label, {
				size: font * 0.85,
				anchor: "end",
				fill: FAINT_TEXT,
				maxChars: 14,
			});
		} else {
			text(svg, plot.x + offsetOf(i), plot.y + plot.h + font * 1.25, label, {
				size: font * 0.85,
				anchor: "middle",
				fill: FAINT_TEXT,
				maxChars: Math.max(4, Math.floor(spacing / (font * 0.55))),
			});
		}
	});
}

// ----------------------------------------------------------- lines/area

function drawLines(
	svg: SVGElement,
	chart: ChartModel,
	layout: Layout,
	font: number,
	filled: boolean,
): void {
	const { plot } = layout;
	const axes = drawCartesianFrame(svg, chart, layout, font, false);
	const count = Math.max(chart.categories.length, 1);
	const step = count > 1 ? plot.w / (count - 1) : 0;
	const x = (i: number) => (count > 1 ? plot.x + i * step : plot.x + plot.w / 2);

	chart.series.forEach((series, s) => {
		const colour = seriesColour(chart, series, s);
		// A null is a gap in the data, so the line is drawn in runs.
		let run: { x: number; y: number }[] = [];
		const flush = (): void => {
			if (run.length === 0) return;
			if (filled && run.length > 1) {
				const area = document.createElementNS(SVG, "polygon");
				const points = [
					`${run[0].x},${axes.zero}`,
					...run.map((p) => `${p.x},${p.y}`),
					`${run[run.length - 1].x},${axes.zero}`,
				];
				area.setAttribute("points", points.join(" "));
				area.setAttribute("fill", colour);
				area.setAttribute("fill-opacity", "0.35");
				svg.appendChild(area);
			}
			const path = document.createElementNS(SVG, "polyline");
			path.setAttribute("points", run.map((p) => `${p.x},${p.y}`).join(" "));
			path.setAttribute("fill", "none");
			path.setAttribute("stroke", colour);
			path.setAttribute("stroke-width", String(Math.max(font * 0.16, 1.4)));
			path.setAttribute("stroke-linejoin", "round");
			path.setAttribute("stroke-linecap", "round");
			svg.appendChild(path);
			for (const point of run) {
				circle(svg, point.x, point.y, Math.max(font * 0.16, 1.6), colour);
			}
			run = [];
		};

		for (let i = 0; i < count; i++) {
			const value = series.values[i];
			if (value === null || value === undefined) {
				flush();
				continue;
			}
			run.push({ x: x(i), y: axes.toValue(value) });
		}
		flush();
	});

	drawCategoryLabels(svg, chart, layout, font, false, count > 1 ? step : plot.w, (i) =>
		count > 1 ? i * step : plot.w / 2,
	);
}

// -------------------------------------------------------------- scatter

function drawScatter(svg: SVGElement, chart: ChartModel, layout: Layout, font: number): void {
	// Scatter x values arrive as category strings; when they are numbers, plot
	// them properly, and otherwise fall back to evenly spaced points.
	const xs = chart.categories.map((c) => Number(c));
	if (xs.some((v) => !Number.isFinite(v))) {
		drawLines(svg, chart, layout, font, false);
		return;
	}

	const { plot } = layout;
	const axes = drawCartesianFrame(svg, chart, layout, font, false);
	const xMin = Math.min(...xs);
	const xMax = Math.max(...xs);
	const xSpan = xMax - xMin || 1;

	chart.series.forEach((series, s) => {
		const colour = seriesColour(chart, series, s);
		series.values.forEach((value, i) => {
			if (value === null || xs[i] === undefined) return;
			const px = plot.x + ((xs[i] - xMin) / xSpan) * plot.w;
			circle(svg, px, axes.toValue(value), Math.max(font * 0.22, 2), colour);
		});
	});
}

// ------------------------------------------------------------ pie/donut

function drawPie(svg: SVGElement, chart: ChartModel, layout: Layout, font: number): void {
	const { plot } = layout;
	const series = chart.series[0];
	if (!series) return;

	const values = series.values.map((v) => (v === null ? 0 : Math.abs(v)));
	const total = values.reduce((sum, v) => sum + v, 0);
	if (total === 0) return;

	const cx = plot.x + plot.w / 2;
	const cy = plot.y + plot.h / 2;
	const radius = Math.min(plot.w, plot.h) / 2;
	const inner = chart.kind === "doughnut" ? radius * clamp(chart.holeSize, 0.1, 0.85) : 0;

	let angle = -Math.PI / 2;
	values.forEach((value, i) => {
		if (value === 0) return;
		const sweep = (value / total) * Math.PI * 2;
		const colour =
			series.pointColors[i] ??
			(chart.palette.length > 0
				? chart.palette[i % chart.palette.length]
				: seriesColour(chart, series, i));
		svg.appendChild(arc(cx, cy, radius, inner, angle, angle + sweep, colour));
		angle += sweep;
	});

	if (inner > 0 && chart.title === "") {
		text(svg, cx, cy + font * 0.35, "", { size: font, anchor: "middle", fill: FAINT_TEXT });
	}
}

/** A pie or doughnut wedge as a path, so one routine covers both. */
function arc(
	cx: number,
	cy: number,
	outer: number,
	inner: number,
	from: number,
	to: number,
	colour: string,
): SVGElement {
	const path = document.createElementNS(SVG, "path");
	const large = to - from > Math.PI ? 1 : 0;
	const p = (r: number, a: number) => `${cx + r * Math.cos(a)},${cy + r * Math.sin(a)}`;

	// A full circle cannot be drawn as one arc; nudge the end so it closes.
	const end = to - from >= Math.PI * 2 ? to - 0.0001 : to;
	const d =
		inner > 0
			? `M ${p(outer, from)} A ${outer} ${outer} 0 ${large} 1 ${p(outer, end)} ` +
				`L ${p(inner, end)} A ${inner} ${inner} 0 ${large} 0 ${p(inner, from)} Z`
			: `M ${cx},${cy} L ${p(outer, from)} A ${outer} ${outer} 0 ${large} 1 ${p(outer, end)} Z`;

	path.setAttribute("d", d);
	path.setAttribute("fill", colour);
	path.setAttribute("stroke", "rgba(255,255,255,0.85)");
	path.setAttribute("stroke-width", "1");
	return path;
}

// --------------------------------------------------------------- legend

function drawLegend(svg: SVGElement, chart: ChartModel, layout: Layout, font: number): void {
	const box = layout.legend;
	if (!box) return;
	const entries =
		chart.kind === "pie" || chart.kind === "doughnut"
			? chart.categories.map((name, i) => ({
					name,
					colour:
						chart.series[0]?.pointColors[i] ??
						(chart.palette.length > 0 ? chart.palette[i % chart.palette.length] : "#4472c4"),
				}))
			: chart.series.map((series, i) => ({ name: series.name, colour: seriesColour(chart, series, i) }));
	if (entries.length === 0) return;

	const swatch = font * 0.8;
	const vertical = chart.legend === "l" || chart.legend === "r";

	if (vertical) {
		entries.forEach((entry, i) => {
			const y = box.y + font * 1.4 * i;
			if (y > box.y + box.h) return;
			square(svg, box.x, y, swatch, entry.colour);
			text(svg, box.x + swatch * 1.5, y + swatch * 0.85, entry.name, {
				size: font * 0.85,
				fill: FAINT_TEXT,
				maxChars: Math.max(4, Math.floor(box.w / (font * 0.55))),
			});
		});
		return;
	}

	// Horizontal legends centre themselves on an estimate of their own width.
	const widths = entries.map((e) => swatch * 1.6 + e.name.length * font * 0.52 + font);
	const total = widths.reduce((sum, w) => sum + w, 0);
	let x = box.x + Math.max(0, (box.w - total) / 2);
	const y = box.y + box.h / 2 - swatch / 2;

	entries.forEach((entry, i) => {
		if (x + widths[i] > box.x + box.w + font) return;
		square(svg, x, y, swatch, entry.colour);
		text(svg, x + swatch * 1.5, y + swatch * 0.85, entry.name, {
			size: font * 0.85,
			fill: FAINT_TEXT,
		});
		x += widths[i];
	});
}

// -------------------------------------------------------------- drawing

function line(svg: SVGElement, x1: number, y1: number, x2: number, y2: number, stroke: string): void {
	const el = document.createElementNS(SVG, "line");
	el.setAttribute("x1", String(x1));
	el.setAttribute("y1", String(y1));
	el.setAttribute("x2", String(x2));
	el.setAttribute("y2", String(y2));
	el.setAttribute("stroke", stroke);
	el.setAttribute("stroke-width", "1");
	svg.appendChild(el);
}

function circle(svg: SVGElement, cx: number, cy: number, r: number, fill: string): void {
	const el = document.createElementNS(SVG, "circle");
	el.setAttribute("cx", String(cx));
	el.setAttribute("cy", String(cy));
	el.setAttribute("r", String(r));
	el.setAttribute("fill", fill);
	svg.appendChild(el);
}

function square(svg: SVGElement, x: number, y: number, size: number, fill: string): void {
	const el = document.createElementNS(SVG, "rect");
	el.setAttribute("x", String(x));
	el.setAttribute("y", String(y));
	el.setAttribute("width", String(size));
	el.setAttribute("height", String(size));
	el.setAttribute("rx", String(size * 0.2));
	el.setAttribute("fill", fill);
	svg.appendChild(el);
}

interface TextOptions {
	size: number;
	anchor?: "start" | "middle" | "end";
	fill?: string;
	weight?: string;
	/** Ellipsise past this many characters, so labels never overrun the plot. */
	maxChars?: number;
}

function text(svg: SVGElement, x: number, y: number, value: string, options: TextOptions): void {
	if (!value) return;
	const el = document.createElementNS(SVG, "text");
	el.setAttribute("x", String(x));
	el.setAttribute("y", String(y));
	el.setAttribute("font-size", String(options.size));
	el.setAttribute("fill", options.fill ?? TEXT_COLOUR);
	if (options.anchor) el.setAttribute("text-anchor", options.anchor);
	if (options.weight) el.setAttribute("font-weight", options.weight);
	el.textContent =
		options.maxChars && value.length > options.maxChars
			? `${value.slice(0, Math.max(1, options.maxChars - 1))}…`
			: value;
	svg.appendChild(el);
}

function clamp(value: number, min: number, max: number): number {
	return Math.max(min, Math.min(max, value));
}
