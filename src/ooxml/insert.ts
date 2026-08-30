import { EMU_PER_PX } from "../pptx/types";
import type { PptxPackage } from "../pptx/package";
import { CT, ensureDefault } from "./contentTypes";
import { REL_BASE, ensureRelationship } from "./rels";
import { A_NS, a, nextShapeId, p as pEl } from "./tree";

const emu = (value: number) => String(Math.round(value * EMU_PER_PX));

export interface NewFrame {
	x: number;
	y: number;
	w: number;
	h: number;
}

/**
 * Builders for the shapes the Insert tab can add.
 *
 * Each produces the markup PowerPoint itself writes, including the `p:style`
 * block that ties a shape to the theme. Without that block a new shape ignores
 * the deck's colours and lands as a bare grey box that looks nothing like the
 * rest of the slide.
 */
function themeStyle(doc: Document, accent = "accent1"): Element {
	const style = pEl(doc, "style");

	const lnRef = a(doc, "lnRef");
	lnRef.setAttribute("idx", "2");
	const lnClr = a(doc, "schemeClr");
	lnClr.setAttribute("val", accent);
	const shade = a(doc, "shade");
	shade.setAttribute("val", "50000");
	lnClr.appendChild(shade);
	lnRef.appendChild(lnClr);

	const fillRef = a(doc, "fillRef");
	fillRef.setAttribute("idx", "1");
	const fillClr = a(doc, "schemeClr");
	fillClr.setAttribute("val", accent);
	fillRef.appendChild(fillClr);

	const effectRef = a(doc, "effectRef");
	effectRef.setAttribute("idx", "0");
	const effectClr = a(doc, "schemeClr");
	effectClr.setAttribute("val", accent);
	effectRef.appendChild(effectClr);

	const fontRef = a(doc, "fontRef");
	fontRef.setAttribute("idx", "minor");
	const fontClr = a(doc, "schemeClr");
	fontClr.setAttribute("val", "lt1");
	fontRef.appendChild(fontClr);

	style.append(lnRef, fillRef, effectRef, fontRef);
	return style;
}

function xfrmElement(doc: Document, frame: NewFrame): Element {
	const xfrm = a(doc, "xfrm");
	const off = a(doc, "off");
	off.setAttribute("x", emu(frame.x));
	off.setAttribute("y", emu(frame.y));
	const ext = a(doc, "ext");
	ext.setAttribute("cx", emu(frame.w));
	ext.setAttribute("cy", emu(frame.h));
	xfrm.append(off, ext);
	return xfrm;
}

function textBody(
	doc: Document,
	options: { text?: string; anchor?: string; align?: string; sizePt?: number; color?: string; wrap?: boolean },
): Element {
	const txBody = pEl(doc, "txBody");
	const bodyPr = a(doc, "bodyPr");
	if (options.anchor) bodyPr.setAttribute("anchor", options.anchor);
	if (options.wrap === false) bodyPr.setAttribute("wrap", "none");
	bodyPr.setAttribute("rtlCol", "0");
	txBody.appendChild(bodyPr);
	txBody.appendChild(a(doc, "lstStyle"));

	const para = a(doc, "p");
	const pPr = a(doc, "pPr");
	if (options.align) pPr.setAttribute("algn", options.align);
	pPr.appendChild(a(doc, "buNone"));
	para.appendChild(pPr);

	const makeProps = (name: string): Element => {
		const props = a(doc, name);
		props.setAttribute("lang", "en-US");
		if (options.sizePt) props.setAttribute("sz", String(Math.round(options.sizePt * 100)));
		if (options.color) {
			const fill = a(doc, "solidFill");
			const clr = a(doc, "srgbClr");
			clr.setAttribute("val", options.color.replace(/^#/, "").toUpperCase());
			fill.appendChild(clr);
			props.appendChild(fill);
		}
		return props;
	};

	if (options.text) {
		const run = a(doc, "r");
		run.appendChild(makeProps("rPr"));
		const t = a(doc, "t");
		t.textContent = options.text;
		run.appendChild(t);
		para.appendChild(run);
	}
	para.appendChild(makeProps("endParaRPr"));
	txBody.appendChild(para);
	return txBody;
}

function nvSpPr(doc: Document, id: number, name: string, isTextBox: boolean): Element {
	const nv = pEl(doc, "nvSpPr");
	const cNvPr = pEl(doc, "cNvPr");
	cNvPr.setAttribute("id", String(id));
	cNvPr.setAttribute("name", name);
	const cNvSpPr = pEl(doc, "cNvSpPr");
	if (isTextBox) cNvSpPr.setAttribute("txBox", "1");
	nv.append(cNvPr, cNvSpPr, pEl(doc, "nvPr"));
	return nv;
}

/** A filled autoshape with the deck's accent colour. */
export function buildShape(
	spTree: Element,
	preset: string,
	frame: NewFrame,
	label: string,
): Element {
	const doc = spTree.ownerDocument;
	const id = nextShapeId(spTree);
	const sp = pEl(doc, "sp");
	sp.appendChild(nvSpPr(doc, id, `${label} ${id}`, false));

	const spPr = pEl(doc, "spPr");
	spPr.appendChild(xfrmElement(doc, frame));
	const geom = a(doc, "prstGeom");
	geom.setAttribute("prst", preset);
	geom.appendChild(a(doc, "avLst"));
	spPr.appendChild(geom);
	sp.appendChild(spPr);

	sp.appendChild(themeStyle(doc));
	sp.appendChild(textBody(doc, { anchor: "ctr", align: "ctr" }));
	return sp;
}

/** A transparent text box: no fill, no outline, text anchored to the top. */
export function buildTextBox(spTree: Element, frame: NewFrame, text = ""): Element {
	const doc = spTree.ownerDocument;
	const id = nextShapeId(spTree);
	const sp = pEl(doc, "sp");
	sp.appendChild(nvSpPr(doc, id, `TextBox ${id}`, true));

	const spPr = pEl(doc, "spPr");
	spPr.appendChild(xfrmElement(doc, frame));
	const geom = a(doc, "prstGeom");
	geom.setAttribute("prst", "rect");
	geom.appendChild(a(doc, "avLst"));
	spPr.appendChild(geom);
	spPr.appendChild(a(doc, "noFill"));
	sp.appendChild(spPr);

	sp.appendChild(textBody(doc, { text, sizePt: 18, anchor: "t" }));
	return sp;
}

/** A straight connector, which is how PowerPoint stores a plain line. */
export function buildLine(spTree: Element, frame: NewFrame): Element {
	const doc = spTree.ownerDocument;
	const id = nextShapeId(spTree);
	const cxn = pEl(doc, "cxnSp");
	const nv = pEl(doc, "nvCxnSpPr");
	const cNvPr = pEl(doc, "cNvPr");
	cNvPr.setAttribute("id", String(id));
	cNvPr.setAttribute("name", `Line ${id}`);
	nv.append(cNvPr, pEl(doc, "cNvCxnSpPr"), pEl(doc, "nvPr"));
	cxn.appendChild(nv);

	const spPr = pEl(doc, "spPr");
	spPr.appendChild(xfrmElement(doc, frame));
	const geom = a(doc, "prstGeom");
	geom.setAttribute("prst", "line");
	geom.appendChild(a(doc, "avLst"));
	spPr.appendChild(geom);
	cxn.appendChild(spPr);
	cxn.appendChild(themeStyle(doc));
	return cxn;
}

/**
 * Add an image to the package and build the picture that shows it.
 *
 * The bytes become a new `ppt/media` part, the extension is registered in
 * `[Content_Types].xml`, and the slide gains a relationship to it — all three
 * are required, and missing any one produces a file PowerPoint offers to repair.
 */
/** The media part path a new picture will take, known before it is created. */
export function nextMediaPath(pkg: PptxPackage, extension: string): string {
	const ext = extension.toLowerCase().replace(/^\./, "");
	const existing = new Set(pkg.partPaths());
	let index = 1;
	while (existing.has(`ppt/media/image${index}.${ext}`)) index++;
	return `ppt/media/image${index}.${ext}`;
}

export function buildPicture(
	pkg: PptxPackage,
	spTree: Element,
	slidePath: string,
	image: { bytes: Uint8Array; extension: string; name: string },
	frame: NewFrame,
	mediaPath = nextMediaPath(pkg, image.extension),
): Element {
	const ext = image.extension.toLowerCase().replace(/^\./, "");
	ensureDefault(pkg, ext, CT.image[ext] ?? "application/octet-stream");
	pkg.addPart(mediaPath, image.bytes);

	const relId = ensureRelationship(pkg, slidePath, `${REL_BASE}/image`, mediaPath);

	const doc = spTree.ownerDocument;
	const id = nextShapeId(spTree);
	const pic = pEl(doc, "pic");

	const nv = pEl(doc, "nvPicPr");
	const cNvPr = pEl(doc, "cNvPr");
	cNvPr.setAttribute("id", String(id));
	cNvPr.setAttribute("name", image.name);
	cNvPr.setAttribute("descr", image.name);
	const cNvPicPr = pEl(doc, "cNvPicPr");
	const locks = a(doc, "picLocks");
	locks.setAttribute("noChangeAspect", "1");
	cNvPicPr.appendChild(locks);
	nv.append(cNvPr, cNvPicPr, pEl(doc, "nvPr"));
	pic.appendChild(nv);

	const blipFill = pEl(doc, "blipFill");
	const blip = a(doc, "blip");
	blip.setAttributeNS(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		"r:embed",
		relId,
	);
	const stretch = a(doc, "stretch");
	stretch.appendChild(a(doc, "fillRect"));
	blipFill.append(blip, stretch);
	pic.appendChild(blipFill);

	const spPr = pEl(doc, "spPr");
	spPr.appendChild(xfrmElement(doc, frame));
	const geom = a(doc, "prstGeom");
	geom.setAttribute("prst", "rect");
	geom.appendChild(a(doc, "avLst"));
	spPr.appendChild(geom);
	pic.appendChild(spPr);

	return pic;
}

/** A table with a filled header row, matching the look PowerPoint inserts. */
export function buildTable(
	spTree: Element,
	frame: NewFrame,
	rows: number,
	columns: number,
): Element {
	const doc = spTree.ownerDocument;
	const id = nextShapeId(spTree);
	const frameEl = pEl(doc, "graphicFrame");

	const nv = pEl(doc, "nvGraphicFramePr");
	const cNvPr = pEl(doc, "cNvPr");
	cNvPr.setAttribute("id", String(id));
	cNvPr.setAttribute("name", `Table ${id}`);
	nv.append(cNvPr, pEl(doc, "cNvGraphicFramePr"), pEl(doc, "nvPr"));
	frameEl.appendChild(nv);

	const xfrm = pEl(doc, "xfrm");
	const off = a(doc, "off");
	off.setAttribute("x", emu(frame.x));
	off.setAttribute("y", emu(frame.y));
	const ext = a(doc, "ext");
	ext.setAttribute("cx", emu(frame.w));
	ext.setAttribute("cy", emu(frame.h));
	xfrm.append(off, ext);
	frameEl.appendChild(xfrm);

	const graphic = a(doc, "graphic");
	const graphicData = a(doc, "graphicData");
	graphicData.setAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/table");

	const tbl = a(doc, "tbl");
	const tblPr = a(doc, "tblPr");
	tblPr.setAttribute("firstRow", "1");
	tblPr.setAttribute("bandRow", "1");
	tbl.appendChild(tblPr);

	const grid = a(doc, "tblGrid");
	const colWidth = frame.w / columns;
	for (let c = 0; c < columns; c++) {
		const col = a(doc, "gridCol");
		col.setAttribute("w", emu(colWidth));
		grid.appendChild(col);
	}
	tbl.appendChild(grid);

	const rowHeight = frame.h / rows;
	for (let r = 0; r < rows; r++) {
		const tr = a(doc, "tr");
		tr.setAttribute("h", emu(rowHeight));
		for (let c = 0; c < columns; c++) {
			const tc = a(doc, "tc");
			const body = textBody(doc, {
				anchor: "ctr",
				sizePt: 14,
				color: r === 0 ? "FFFFFF" : undefined,
			});
			// The table's text body is DrawingML, not PresentationML.
			const renamed = doc.createElementNS(A_NS, "a:txBody");
			while (body.firstChild) renamed.appendChild(body.firstChild);
			tc.appendChild(renamed);

			const tcPr = a(doc, "tcPr");
			for (const side of ["lnL", "lnR", "lnT", "lnB"]) {
				const ln = a(doc, side);
				ln.setAttribute("w", "12700");
				const fill = a(doc, "solidFill");
				const clr = a(doc, "srgbClr");
				clr.setAttribute("val", "C9D1DB");
				fill.appendChild(clr);
				ln.appendChild(fill);
				tcPr.appendChild(ln);
			}
			const fill = a(doc, "solidFill");
			const clr = a(doc, "schemeClr");
			clr.setAttribute("val", r === 0 ? "accent1" : "lt1");
			fill.appendChild(clr);
			tcPr.appendChild(fill);
			tc.appendChild(tcPr);
			tr.appendChild(tc);
		}
		tbl.appendChild(tr);
	}

	graphicData.appendChild(tbl);
	graphic.appendChild(graphicData);
	frameEl.appendChild(graphic);
	return frameEl;
}
