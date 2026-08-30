import { geometryCss } from "../pptx/geometry";
import { renderChart } from "./renderChart";
import { t } from "../i18n";
import type {
	ChartShape,
	Deck,
	Fill,
	Frame,
	GroupShape,
	ImageShape,
	LineShape,
	Paragraph,
	Run,
	Shape,
	Slide,
	Stroke,
	TableShape,
	TextBody,
	TextShape,
} from "../pptx/types";

/** Fallbacks appended to every font stack so CJK and symbol runs still resolve. */
const FONT_FALLBACK =
	'"Helvetica Neue", Helvetica, Arial, "Hiragino Sans", "Hiragino Kaku Gothic ProN", "Yu Gothic", Meiryo, "Segoe UI", sans-serif';

/**
 * Rendered element -> model object. The edit layer needs to get from a DOM node
 * the user clicked back to the XML it came from; a WeakMap keeps that link
 * without putting non-serialisable junk on the elements themselves (which would
 * also leak into PNG export).
 */
export const textBodyRegistry = new WeakMap<HTMLElement, TextBody>();
export const paragraphRegistry = new WeakMap<HTMLElement, Paragraph>();
export const runRegistry = new WeakMap<HTMLElement, Run>();
/** Top-level, slide-owned shapes only — the ones that may be moved and resized. */
export const shapeRegistry = new WeakMap<HTMLElement, Shape>();

export interface RenderOptions {
	/** Draw a subtle border around the slide. */
	border?: boolean;
}

/**
 * The part whose text may be edited during the current render. Text inherited
 * from a layout or master is shown but not editable: changing it would silently
 * rewrite every slide built on that template.
 *
 * Rendering is synchronous and single-threaded, so a module-scoped value here is
 * safe and keeps an extra parameter out of every render function.
 */
let editablePart: string | null = null;

/**
 * Render one slide at its native pixel size. Callers scale the result with a CSS
 * transform, which keeps text crisp and makes zooming free.
 */
export function renderSlide(deck: Deck, slide: Slide, options: RenderOptions = {}): HTMLElement {
	const root = document.createElement("div");
	root.addClass("pptx-slide");
	root.dataset.slideIndex = String(slide.index);
	Object.assign(root.style, {
		position: "relative",
		width: `${deck.width}px`,
		height: `${deck.height}px`,
		overflow: "hidden",
		boxSizing: "border-box",
		transformOrigin: "top left",
	});
	if (options.border) root.style.border = "1px solid rgba(0,0,0,0.12)";
	applyFill(root, slide.background ?? { kind: "solid", color: "#ffffff" });

	editablePart = slide.partPath;
	try {
		for (const shape of slide.shapes) {
			const el = renderShape(shape);
			if (!el) continue;
			// Layout and master artwork is drawn but not selectable: dragging it
			// would move the same logo on every slide using that template.
			if (shape.source && shape.sourcePart === slide.partPath) {
				el.dataset.selectable = "1";
				el.dataset.shapeId = shape.id;
				shapeRegistry.set(el, shape);
			}
			root.appendChild(el);
		}
	} finally {
		editablePart = null;
	}
	return root;
}

function renderShape(shape: Shape): HTMLElement | null {
	if (shape.hidden) return null;
	switch (shape.kind) {
		case "shape":
			return renderTextShape(shape);
		case "image":
			return renderImage(shape);
		case "table":
			return renderTable(shape);
		case "chart":
			return renderChartShape(shape);
		case "group":
			return renderGroup(shape);
		case "line":
			return renderLine(shape);
	}
}

/** Position an element at a frame, applying rotation and flips. */
function positioned(frame: Frame, className: string): HTMLElement {
	const el = document.createElement("div");
	el.addClass(className);
	Object.assign(el.style, {
		position: "absolute",
		left: `${frame.x}px`,
		top: `${frame.y}px`,
		width: `${frame.w}px`,
		height: `${frame.h}px`,
		boxSizing: "border-box",
	});
	const transforms: string[] = [];
	if (frame.rot) transforms.push(`rotate(${frame.rot}deg)`);
	if (frame.flipH) transforms.push("scaleX(-1)");
	if (frame.flipV) transforms.push("scaleY(-1)");
	if (transforms.length) {
		el.style.transform = transforms.join(" ");
		el.style.transformOrigin = "center center";
	}
	return el;
}

function applyFill(el: HTMLElement, fill: Fill | null): void {
	if (!fill) return;
	switch (fill.kind) {
		case "solid":
			el.style.background = fill.color;
			break;
		case "gradient":
			el.style.background = fill.css;
			break;
		case "image":
			if (fill.url) {
				if (fill.mediaPath) el.dataset.mediaPath = fill.mediaPath;
				el.style.backgroundImage = `url("${fill.url}")`;
				el.style.backgroundSize = "cover";
				el.style.backgroundPosition = "center";
				if (fill.opacity < 1) el.style.opacity = String(fill.opacity);
			}
			break;
		case "none":
			el.style.background = "transparent";
			break;
	}
}

function applyStroke(el: HTMLElement, stroke: Stroke | null, clipped: boolean): void {
	if (!stroke) return;
	if (clipped) {
		// A CSS border would be clipped away with the corners, so approximate the
		// outline with a drop-shadow that follows the clip path.
		el.style.filter = `drop-shadow(0 0 ${stroke.width}px ${stroke.color})`;
		return;
	}
	el.style.border = `${stroke.width}px ${stroke.style} ${stroke.color}`;
}

// --------------------------------------------------------------- shapes

function renderTextShape(shape: TextShape): HTMLElement {
	const el = positioned(shape.frame, "pptx-shape");
	if (shape.placeholder) el.dataset.placeholder = shape.placeholder;

	const geo = geometryCss(shape.geom, shape.frame.w, shape.frame.h);
	if (geo.borderRadius) el.style.borderRadius = geo.borderRadius;
	if (geo.clipPath) el.style.clipPath = geo.clipPath;
	applyFill(el, shape.fill);
	applyStroke(el, shape.stroke, geo.clipped);

	if (shape.text) el.appendChild(renderTextBody(shape.text));
	return el;
}

function renderTextBody(body: TextBody): HTMLElement {
	const box = document.createElement("div");
	box.addClass("pptx-text");
	textBodyRegistry.set(box, body);
	if (body.source && body.sourcePart === editablePart) box.dataset.editable = "1";
	Object.assign(box.style, {
		position: "absolute",
		inset: "0",
		display: "flex",
		flexDirection: "column",
		justifyContent:
			body.anchor === "middle" ? "center" : body.anchor === "bottom" ? "flex-end" : "flex-start",
		paddingLeft: `${body.insets[0]}px`,
		paddingTop: `${body.insets[1]}px`,
		paddingRight: `${body.insets[2]}px`,
		paddingBottom: `${body.insets[3]}px`,
		boxSizing: "border-box",
		overflow: "visible",
		whiteSpace: body.wrap ? "pre-wrap" : "pre",
		wordBreak: "break-word",
	});

	if (body.vertical === "vert") {
		box.style.writingMode = "vertical-rl";
	} else if (body.vertical === "vert270") {
		box.style.writingMode = "vertical-rl";
		box.style.transform = "rotate(180deg)";
	}

	for (const para of body.paragraphs) {
		box.appendChild(renderParagraph(para, body));
	}
	return box;
}

function renderParagraph(para: Paragraph, body: TextBody): HTMLElement {
	const scale = body.fontScale;
	const el = document.createElement("div");
	el.addClass("pptx-para");
	Object.assign(el.style, {
		textAlign: para.align,
		marginTop: `${para.spaceBefore * scale}px`,
		marginBottom: `${para.spaceAfter * scale}px`,
		paddingLeft: `${Math.max(0, para.marginLeft)}px`,
	});

	const spacing = para.lineSpacing;
	if (spacing !== null) {
		el.style.lineHeight =
			spacing < 0
				? `${-spacing * scale}px`
				: String(Math.max(0.5, spacing - body.lineSpaceReduction));
	} else {
		el.style.lineHeight = String(Math.max(0.5, 1.2 - body.lineSpaceReduction));
	}

	paragraphRegistry.set(el, para);

	const runsEl = document.createElement("div");
	runsEl.addClass("pptx-runs");
	para.runs.forEach((run, i) => {
		runsEl.appendChild(renderRun(run, scale, i));
	});

	if (!para.bullet) {
		el.appendChild(runsEl);
		return el;
	}

	// Hanging indent: the bullet occupies the negative indent to the left of the text.
	const row = document.createElement("div");
	Object.assign(row.style, {
		display: "flex",
		alignItems: "baseline",
		gap: "0",
	});
	const marker = document.createElement("span");
	marker.addClass("pptx-bullet");
	// The glyph is generated from the list style, not typed, so keep it out of
	// the editable flow: the caret should never land inside a bullet.
	marker.contentEditable = "false";
	marker.style.userSelect = "none";
	const firstSize = (para.runs[0]?.size ?? 18) * scale;
	Object.assign(marker.style, {
		flex: "0 0 auto",
		minWidth: `${Math.max(Math.abs(para.indent), firstSize * 0.9)}px`,
		color: para.bullet.color ?? para.runs[0]?.color ?? "inherit",
		fontSize: `${firstSize * para.bullet.scale}px`,
	});
	if (para.bullet.font) marker.style.fontFamily = `"${para.bullet.font}", ${FONT_FALLBACK}`;
	marker.setText(para.bullet.text);
	runsEl.style.flex = "1 1 auto";
	runsEl.style.minWidth = "0";
	row.appendChild(marker);
	row.appendChild(runsEl);
	el.appendChild(row);
	return el;
}

function renderRun(run: Run, scale: number, index: number): HTMLElement {
	const el = document.createElement(run.link ? "a" : "span");
	el.dataset.run = String(index);
	runRegistry.set(el, run);
	if (run.link && el instanceof HTMLAnchorElement) {
		el.href = run.link;
		el.target = "_blank";
		el.rel = "noopener";
	}
	Object.assign(el.style, {
		fontSize: `${run.size * scale}px`,
		fontWeight: run.bold ? "700" : "400",
		fontStyle: run.italic ? "italic" : "normal",
		color: run.color,
		fontFamily: `"${run.font}", ${FONT_FALLBACK}`,
	});
	const decorations: string[] = [];
	if (run.underline) decorations.push("underline");
	if (run.strike) decorations.push("line-through");
	el.style.textDecoration = decorations.length ? decorations.join(" ") : "none";
	if (run.spacing) el.style.letterSpacing = `${run.spacing * scale}px`;
	if (run.highlight) el.style.background = run.highlight;
	if (run.baseline > 0) el.style.verticalAlign = "super";
	if (run.baseline < 0) el.style.verticalAlign = "sub";
	if (run.baseline !== 0) el.style.fontSize = `${run.size * scale * 0.65}px`;
	el.textContent = run.text;
	return el;
}

function renderImage(shape: ImageShape): HTMLElement {
	const el = positioned(shape.frame, "pptx-image");
	el.style.overflow = "hidden";
	const geo = geometryCss(shape.geom, shape.frame.w, shape.frame.h);
	if (geo.borderRadius) el.style.borderRadius = geo.borderRadius;
	if (geo.clipPath) el.style.clipPath = geo.clipPath;
	applyStroke(el, shape.stroke, geo.clipped);

	if (!shape.url) {
		// EMF/WMF and friends: show a labelled placeholder rather than a blank hole.
		Object.assign(el.style, {
			display: "flex",
			alignItems: "center",
			justifyContent: "center",
			background: "rgba(0,0,0,0.04)",
			border: "1px dashed rgba(0,0,0,0.25)",
			color: "rgba(0,0,0,0.55)",
			fontSize: "12px",
			padding: "4px",
			textAlign: "center",
		});
		el.setText(shape.label || t("render.unsupportedImage"));
		return el;
	}

	const img = document.createElement("img");
	img.src = shape.url;
	img.alt = shape.label;
	if (shape.mediaPath) img.dataset.mediaPath = shape.mediaPath;
	img.draggable = false;
	if (shape.crop) {
		const { l, t, r, b } = shape.crop;
		const wFactor = 1 - l - r;
		const hFactor = 1 - t - b;
		Object.assign(img.style, {
			position: "absolute",
			width: `${100 / Math.max(wFactor, 0.001)}%`,
			height: `${100 / Math.max(hFactor, 0.001)}%`,
			left: `${(-l / Math.max(wFactor, 0.001)) * 100}%`,
			top: `${(-t / Math.max(hFactor, 0.001)) * 100}%`,
			objectFit: "fill",
		});
	} else {
		Object.assign(img.style, { width: "100%", height: "100%", objectFit: "fill" });
	}
	el.appendChild(img);
	return el;
}

function renderTable(shape: TableShape): HTMLElement {
	const el = positioned(shape.frame, "pptx-table");
	const table = document.createElement("table");
	Object.assign(table.style, {
		width: "100%",
		height: "100%",
		borderCollapse: "collapse",
		tableLayout: "fixed",
	});

	const colgroup = document.createElement("colgroup");
	for (const w of shape.table.columns) {
		const col = document.createElement("col");
		col.style.width = `${w}px`;
		colgroup.appendChild(col);
	}
	table.appendChild(colgroup);

	const tbody = document.createElement("tbody");
	shape.table.rows.forEach((row, rowIndex) => {
		const tr = document.createElement("tr");
		if (row.height) tr.style.height = `${row.height}px`;
		row.cells.forEach((cell, colIndex) => {
			if (cell.merged) return;
			const td = document.createElement("td");
			// Grid coordinates, so a click can be turned back into a cell.
			td.dataset.cellRow = String(rowIndex);
			td.dataset.cellCol = String(colIndex);
			if (cell.colSpan > 1) td.colSpan = cell.colSpan;
			if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
			Object.assign(td.style, {
				verticalAlign:
					cell.anchor === "middle" ? "middle" : cell.anchor === "bottom" ? "bottom" : "top",
				padding: `${cell.margins[1]}px ${cell.margins[2]}px ${cell.margins[3]}px ${cell.margins[0]}px`,
				overflow: "hidden",
			});
			applyFill(td, cell.fill);
			for (const [side, stroke] of [
				["Left", cell.borders.left],
				["Top", cell.borders.top],
				["Right", cell.borders.right],
				["Bottom", cell.borders.bottom],
			] as const) {
				if (stroke) {
					td.style.setProperty(
						`border-${side.toLowerCase()}`,
						`${stroke.width}px ${stroke.style} ${stroke.color}`,
					);
				}
			}
			if (cell.text) {
				const body = renderTextBody(cell.text);
				// Inside a table cell the text flows normally rather than filling a box.
				Object.assign(body.style, { position: "static", inset: "auto", padding: "0" });
				td.appendChild(body);
			}
			tr.appendChild(td);
		});
		tbody.appendChild(tr);
	});
	table.appendChild(tbody);
	el.appendChild(table);
	return el;
}

function renderChartShape(shape: ChartShape): HTMLElement {
	const el = positioned(shape.frame, "pptx-chart");
	el.style.overflow = "hidden";
	el.appendChild(renderChart(shape));
	return el;
}

function renderGroup(shape: GroupShape): HTMLElement {
	const el = positioned(shape.frame, "pptx-group");
	const inner = document.createElement("div");
	const sx = shape.childOffset.w > 0 ? shape.frame.w / shape.childOffset.w : 1;
	const sy = shape.childOffset.h > 0 ? shape.frame.h / shape.childOffset.h : 1;
	Object.assign(inner.style, {
		position: "absolute",
		left: "0",
		top: "0",
		width: `${shape.childOffset.w || shape.frame.w}px`,
		height: `${shape.childOffset.h || shape.frame.h}px`,
		transformOrigin: "top left",
		transform: `scale(${sx}, ${sy}) translate(${-shape.childOffset.x}px, ${-shape.childOffset.y}px)`,
	});
	for (const kid of shape.children) {
		const kidEl = renderShape(kid);
		if (kidEl) inner.appendChild(kidEl);
	}
	el.appendChild(inner);
	return el;
}

function renderLine(shape: LineShape): HTMLElement {
	const el = positioned(shape.frame, "pptx-line");
	if (!shape.stroke) return el;
	const w = Math.max(shape.frame.w, 1);
	const h = Math.max(shape.frame.h, 1);
	const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
	svg.setAttribute("width", "100%");
	svg.setAttribute("height", "100%");
	svg.setAttribute("viewBox", `0 0 ${w} ${h}`);
	svg.setAttribute("preserveAspectRatio", "none");
	svg.style.overflow = "visible";

	const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
	line.setAttribute("x1", "0");
	line.setAttribute("y1", "0");
	line.setAttribute("x2", String(w));
	line.setAttribute("y2", String(h));
	line.setAttribute("stroke", shape.stroke.color);
	line.setAttribute("stroke-width", String(shape.stroke.width));
	if (shape.stroke.style === "dashed") {
		line.setAttribute("stroke-dasharray", `${shape.stroke.width * 4} ${shape.stroke.width * 3}`);
	} else if (shape.stroke.style === "dotted") {
		line.setAttribute("stroke-linecap", "round");
		line.setAttribute("stroke-dasharray", `0 ${shape.stroke.width * 2}`);
	}
	svg.appendChild(line);
	el.appendChild(svg);
	return el;
}
