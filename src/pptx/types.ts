/** The renderer-facing model. Parsing turns OOXML into these; rendering only reads these. */

export interface Deck {
	/** Slide size in CSS pixels at 100% zoom. */
	width: number;
	height: number;
	slides: Slide[];
	title: string;
}

export interface Slide {
	/** 1-based, matching what PowerPoint shows. */
	index: number;
	name: string;
	/** Package part path, e.g. "ppt/slides/slide1.xml". */
	partPath: string;
	background: Fill | null;
	/** Layout/master decoration first, then the slide's own shapes. */
	shapes: Shape[];
	/** How many leading entries of `shapes` came from the layout or master. */
	templateShapes: number;
	notes: string;
}

export type Fill =
	| { kind: "solid"; color: string }
	| { kind: "gradient"; css: string }
	| { kind: "image"; url: string | null; mediaPath: string | null; opacity: number }
	| { kind: "none" };

export interface Stroke {
	color: string;
	/** Width in px. */
	width: number;
	/** CSS border-style. */
	style: string;
}

export interface Frame {
	x: number;
	y: number;
	w: number;
	h: number;
	/** Degrees, clockwise. */
	rot: number;
	flipH: boolean;
	flipV: boolean;
}

export interface Crop {
	l: number;
	t: number;
	r: number;
	b: number;
}

export interface Run {
	text: string;
	/** The a:r / a:fld / a:br element this run came from; null for synthesised runs. */
	source: Element | null;
	/** px */
	size: number;
	bold: boolean;
	italic: boolean;
	underline: boolean;
	strike: boolean;
	color: string;
	font: string;
	/** Super/subscript, as a percentage of font size. */
	baseline: number;
	/** Letter spacing in px. */
	spacing: number;
	link: string | null;
	highlight: string | null;
}

export interface Bullet {
	kind: "char" | "number" | "none";
	/** Rendered glyph or number label. */
	text: string;
	color: string | null;
	font: string | null;
	/** Size relative to the first run, as a fraction. */
	scale: number;
}

export interface Paragraph {
	/** The a:p element this paragraph came from. */
	source: Element | null;
	level: number;
	align: "left" | "center" | "right" | "justify";
	bullet: Bullet | null;
	/** px */
	marginLeft: number;
	indent: number;
	spaceBefore: number;
	spaceAfter: number;
	/** Multiplier (1.0 == single) or null to inherit. */
	lineSpacing: number | null;
	runs: Run[];
}

export interface TextBody {
	/** The a:txBody element, and the package part it lives in. */
	source: Element | null;
	sourcePart: string;
	anchor: "top" | "middle" | "bottom";
	/** Insets in px: left, top, right, bottom. */
	insets: [number, number, number, number];
	wrap: boolean;
	/** normAutofit shrink factors, already normalised to fractions. */
	fontScale: number;
	lineSpaceReduction: number;
	paragraphs: Paragraph[];
	/** Rotate the text block inside its shape (vert / vert270). */
	vertical: "horz" | "vert" | "vert270";
}

export interface TableCell {
	text: TextBody | null;
	fill: Fill | null;
	borders: {
		left: Stroke | null;
		top: Stroke | null;
		right: Stroke | null;
		bottom: Stroke | null;
	};
	colSpan: number;
	rowSpan: number;
	/** True when this cell is covered by another cell's span. */
	merged: boolean;
	anchor: "top" | "middle" | "bottom";
	margins: [number, number, number, number];
}

export interface Table {
	/** Column widths in px. */
	columns: number[];
	rows: { height: number; cells: TableCell[] }[];
}

interface ShapeBase {
	id: string;
	name: string;
	frame: Frame;
	hidden: boolean;
	/** The XML element this shape came from, for write-back. */
	source: Element | null;
	/** The package part `source` lives in. */
	sourcePart: string;
}

export interface TextShape extends ShapeBase {
	kind: "shape";
	/** a:prstGeom preset name, e.g. "roundRect". */
	geom: string;
	fill: Fill | null;
	stroke: Stroke | null;
	text: TextBody | null;
	/** Set for placeholders, so the renderer can tag them for styling. */
	placeholder: string | null;
}

export interface ImageShape extends ShapeBase {
	kind: "image";
	url: string | null;
	mediaPath: string | null;
	crop: Crop | null;
	/** Alt text, also used as the placeholder label for undecodable formats. */
	label: string;
	geom: string;
	stroke: Stroke | null;
}

export interface TableShape extends ShapeBase {
	kind: "table";
	table: Table;
}

export interface ChartShape extends ShapeBase {
	kind: "chart";
	chartType: string;
	title: string;
	series: string[];
	categories: string[];
}

export interface GroupShape extends ShapeBase {
	kind: "group";
	/** The group's child coordinate space (a:chOff / a:chExt), in px. */
	childOffset: { x: number; y: number; w: number; h: number };
	children: Shape[];
}

export interface LineShape extends ShapeBase {
	kind: "line";
	stroke: Stroke | null;
	/** Arrowheads, if any. */
	headArrow: boolean;
	tailArrow: boolean;
}

export type Shape =
	| TextShape
	| ImageShape
	| TableShape
	| ChartShape
	| GroupShape
	| LineShape;

export const EMU_PER_PX = 9525;
export const PT_TO_PX = 96 / 72;

export function emuToPx(emu: number): number {
	return emu / EMU_PER_PX;
}
