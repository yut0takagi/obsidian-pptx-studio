import { EMU_PER_PX } from "../pptx/types";
import { child, children } from "../pptx/xml";
import { A_NS } from "./tree";

/**
 * Row and column surgery on a:tbl.
 *
 * OOXML tables are a dense grid: every position has an a:tc, and a merged cell
 * is expressed by the top-left cell carrying gridSpan/rowSpan while the covered
 * positions stay in the document marked hMerge/vMerge. Keeping that invariant is
 * what every operation here is really about — a table with a missing a:tc opens
 * as "repair needed".
 */

export function tableOf(shapeSource: Element): Element | null {
	const graphicData = child(child(shapeSource, "graphic"), "graphicData");
	return child(graphicData, "tbl");
}

export function tableRows(tbl: Element): Element[] {
	return children(tbl, "tr");
}

export function rowCells(tr: Element): Element[] {
	return children(tr, "tc");
}

export function gridColumns(tbl: Element): Element[] {
	return children(child(tbl, "tblGrid"), "gridCol");
}

/** An empty cell that copies a neighbour's borders and fill. */
function blankCell(doc: Document, template: Element | null): Element {
	const tc = doc.createElementNS(A_NS, "a:tc");
	const txBody = doc.createElementNS(A_NS, "a:txBody");
	txBody.appendChild(doc.createElementNS(A_NS, "a:bodyPr"));
	txBody.appendChild(doc.createElementNS(A_NS, "a:lstStyle"));
	const para = doc.createElementNS(A_NS, "a:p");
	// Carry the neighbour's end-of-paragraph run properties so typing into the new
	// cell comes out in the table's font rather than a bare default.
	const templateEnd = template ? child(child(child(template, "txBody"), "p"), "endParaRPr") : null;
	if (templateEnd) para.appendChild(templateEnd.cloneNode(true));
	txBody.appendChild(para);
	tc.appendChild(txBody);

	const templatePr = template ? child(template, "tcPr") : null;
	tc.appendChild(
		templatePr
			? (templatePr.cloneNode(true) as Element)
			: doc.createElementNS(A_NS, "a:tcPr"),
	);
	return tc;
}

export function insertRow(tbl: Element, at: number, copyFrom: number): boolean {
	const rows = tableRows(tbl);
	const template = rows[Math.max(0, Math.min(rows.length - 1, copyFrom))];
	if (!template) return false;
	const doc = tbl.ownerDocument;

	const tr = doc.createElementNS(A_NS, "a:tr");
	const height = template.getAttribute("h");
	if (height) tr.setAttribute("h", height);
	for (const cell of rowCells(template)) {
		const copy = blankCell(doc, cell);
		// A new row cannot continue a vertical merge from above.
		const span = cell.getAttribute("gridSpan");
		if (span) copy.setAttribute("gridSpan", span);
		if (cell.getAttribute("hMerge")) copy.setAttribute("hMerge", "1");
		tr.appendChild(copy);
	}

	const anchor = rows[at] ?? null;
	tbl.insertBefore(tr, anchor);
	return true;
}

export function deleteRow(tbl: Element, index: number): boolean {
	const rows = tableRows(tbl);
	if (rows.length <= 1 || !rows[index]) return false;
	tbl.removeChild(rows[index]);
	return true;
}

export function insertColumn(tbl: Element, at: number, copyFrom: number): boolean {
	const grid = child(tbl, "tblGrid");
	const columns = gridColumns(tbl);
	const template = columns[Math.max(0, Math.min(columns.length - 1, copyFrom))];
	if (!grid || !template) return false;
	const doc = tbl.ownerDocument;

	// Take the new column's width out of the one it was copied from, so the table
	// keeps its overall width instead of growing off the slide.
	const width = Number(template.getAttribute("w") ?? 0);
	const half = Math.max(Math.round(width / 2), Math.round(0.25 * EMU_PER_PX));
	template.setAttribute("w", String(Math.max(half, width - half)));

	const col = doc.createElementNS(A_NS, "a:gridCol");
	col.setAttribute("w", String(half));
	grid.insertBefore(col, columns[at] ?? null);

	for (const tr of tableRows(tbl)) {
		const cells = rowCells(tr);
		const neighbour = cells[Math.max(0, Math.min(cells.length - 1, copyFrom))] ?? null;
		tr.insertBefore(blankCell(doc, neighbour), cells[at] ?? null);
	}
	return true;
}

export function deleteColumn(tbl: Element, index: number): boolean {
	const grid = child(tbl, "tblGrid");
	const columns = gridColumns(tbl);
	if (!grid || columns.length <= 1 || !columns[index]) return false;

	// Give the removed width back to a neighbour rather than shrinking the table.
	const removed = Number(columns[index].getAttribute("w") ?? 0);
	const neighbour = columns[index + 1] ?? columns[index - 1];
	if (neighbour) {
		neighbour.setAttribute("w", String(Number(neighbour.getAttribute("w") ?? 0) + removed));
	}
	grid.removeChild(columns[index]);

	for (const tr of tableRows(tbl)) {
		const cells = rowCells(tr);
		if (cells[index]) tr.removeChild(cells[index]);
	}
	return true;
}

export interface CellRange {
	r0: number;
	c0: number;
	r1: number;
	c1: number;
}

export function normaliseRange(range: CellRange): CellRange {
	return {
		r0: Math.min(range.r0, range.r1),
		c0: Math.min(range.c0, range.c1),
		r1: Math.max(range.r0, range.r1),
		c1: Math.max(range.c0, range.c1),
	};
}

/**
 * Merge a rectangle of cells into the one at its top-left.
 *
 * The covered cells stay in the document — that is how OOXML represents a merge —
 * but are flagged, and their text is folded into the surviving cell so nothing
 * the user typed silently disappears.
 */
export function mergeCells(tbl: Element, range: CellRange): boolean {
	const box = normaliseRange(range);
	if (box.r0 === box.r1 && box.c0 === box.c1) return false;
	const rows = tableRows(tbl);
	const anchor = rowCells(rows[box.r0] ?? tbl)[box.c0];
	if (!anchor) return false;

	const collected: string[] = [];
	for (let r = box.r0; r <= box.r1; r++) {
		const cells = rowCells(rows[r] ?? tbl);
		for (let c = box.c0; c <= box.c1; c++) {
			const tc = cells[c];
			if (!tc) continue;
			if (tc !== anchor) {
				const text = (child(tc, "txBody")?.textContent ?? "").trim();
				if (text) collected.push(text);
				clearCellText(tc);
			}
			if (r > box.r0) tc.setAttribute("vMerge", "1");
			if (c > box.c0) tc.setAttribute("hMerge", "1");
		}
	}

	anchor.setAttribute("gridSpan", String(box.c1 - box.c0 + 1));
	anchor.setAttribute("rowSpan", String(box.r1 - box.r0 + 1));
	if (collected.length > 0) appendCellText(anchor, collected);
	return true;
}

/** Undo a merge, leaving every covered position an ordinary empty cell. */
export function splitCells(tbl: Element, range: CellRange): boolean {
	const box = normaliseRange(range);
	const rows = tableRows(tbl);
	let changed = false;
	for (let r = box.r0; r <= box.r1; r++) {
		const cells = rowCells(rows[r] ?? tbl);
		for (let c = box.c0; c <= box.c1; c++) {
			const tc = cells[c];
			if (!tc) continue;
			const span = Number(tc.getAttribute("gridSpan") ?? 1);
			const rowSpan = Number(tc.getAttribute("rowSpan") ?? 1);
			if (span > 1 || rowSpan > 1 || tc.getAttribute("hMerge") || tc.getAttribute("vMerge")) {
				changed = true;
			}
			tc.removeAttribute("gridSpan");
			tc.removeAttribute("rowSpan");
			tc.removeAttribute("hMerge");
			tc.removeAttribute("vMerge");
		}
	}
	return changed;
}

/** The a:tc at a grid position, or null. */
export function cellAt(tbl: Element, row: number, column: number): Element | null {
	const tr = tableRows(tbl)[row];
	if (!tr) return null;
	return rowCells(tr)[column] ?? null;
}

function clearCellText(tc: Element): void {
	const txBody = child(tc, "txBody");
	if (!txBody) return;
	for (const para of children(txBody, "p")) {
		for (const run of [...children(para, "r"), ...children(para, "fld"), ...children(para, "br")]) {
			para.removeChild(run);
		}
	}
}

function appendCellText(tc: Element, lines: string[]): void {
	const txBody = child(tc, "txBody");
	if (!txBody) return;
	const doc = tc.ownerDocument;
	const template = children(txBody, "p")[0] ?? null;
	for (const line of lines) {
		const para = template
			? (template.cloneNode(true) as Element)
			: doc.createElementNS(A_NS, "a:p");
		for (const run of [...children(para, "r"), ...children(para, "fld"), ...children(para, "br")]) {
			para.removeChild(run);
		}
		const r = doc.createElementNS(A_NS, "a:r");
		const t = doc.createElementNS(A_NS, "a:t");
		t.textContent = line;
		r.appendChild(t);
		para.appendChild(r);
		txBody.appendChild(para);
	}
}
