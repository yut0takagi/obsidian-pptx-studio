import { setOrderedChild } from "../ooxml/format";
import { relsPathFor } from "../ooxml/rels";
import {
	type CellRange,
	cellAt,
	deleteColumn,
	deleteRow,
	gridColumns,
	insertColumn,
	insertRow,
	mergeCells,
	normaliseRange,
	splitCells,
	tableOf,
	tableRows,
} from "../ooxml/table";
import { A_NS } from "../ooxml/tree";
import type { Shape } from "../pptx/types";
import type { CommandContext } from "./commands";
import { selectedShapes } from "./commands";

const TCPR_ORDER = [
	"lnL",
	"lnR",
	"lnT",
	"lnB",
	"lnTlToBr",
	"lnBlToTr",
	"cell3D",
	"noFill",
	"solidFill",
	"gradFill",
	"blipFill",
	"pattFill",
	"grpFill",
	"headers",
	"extLst",
];

/**
 * Which cells of which table are selected.
 *
 * Table editing needs a second, finer selection than the shape one: the shape
 * selection says "this table", and this says "these cells inside it". Keeping
 * them apart means clicking a table still behaves like clicking a shape.
 */
export class TableSelection {
	private shape: string | null = null;
	private range: CellRange | null = null;
	private readonly listeners = new Set<() => void>();

	get shapeId(): string | null {
		return this.shape;
	}

	get cells(): CellRange | null {
		return this.range ? normaliseRange(this.range) : null;
	}

	get isEmpty(): boolean {
		return this.range === null;
	}

	/** Click a cell; `extend` grows the rectangle from the existing anchor. */
	select(shapeId: string, row: number, column: number, extend: boolean): void {
		if (extend && this.shape === shapeId && this.range) {
			this.range = { ...this.range, r1: row, c1: column };
		} else {
			this.shape = shapeId;
			this.range = { r0: row, c0: column, r1: row, c1: column };
		}
		this.notify();
	}

	clear(): void {
		if (this.shape === null && this.range === null) return;
		this.shape = null;
		this.range = null;
		this.notify();
	}

	/** Forget the selection when it no longer refers to the selected shape. */
	retain(shapeIds: ReadonlySet<string>): void {
		if (this.shape && !shapeIds.has(this.shape)) this.clear();
	}

	onChange(listener: () => void): () => void {
		this.listeners.add(listener);
		return () => this.listeners.delete(listener);
	}

	private notify(): void {
		for (const listener of this.listeners) listener();
	}

	/** Mark the selected cells in a rendered slide. */
	paint(slideEl: HTMLElement | null): void {
		if (!slideEl) return;
		for (const el of Array.from(slideEl.querySelectorAll(".pptx-cell-selected"))) {
			el.removeClass("pptx-cell-selected");
		}
		const box = this.cells;
		if (!this.shape || !box) return;
		for (const el of Array.from(slideEl.querySelectorAll<HTMLElement>("[data-shape-id]"))) {
			if (el.dataset.shapeId !== this.shape) continue;
			for (const td of Array.from(el.querySelectorAll<HTMLElement>("[data-cell-row]"))) {
				const r = Number(td.dataset.cellRow);
				const c = Number(td.dataset.cellCol);
				if (r >= box.r0 && r <= box.r1 && c >= box.c0 && c <= box.c1) {
					td.addClass("pptx-cell-selected");
				}
			}
		}
	}
}

interface TableTarget {
	shape: Shape;
	tbl: Element;
	range: CellRange;
}

/** The table and cells a table command should act on. */
export function tableTarget(ctx: CommandContext | null, sel: TableSelection): TableTarget | null {
	if (!ctx) return null;
	const shape = selectedShapes(ctx).find((s) => s.kind === "table" && s.id === sel.shapeId);
	if (!shape?.source) return null;
	const tbl = tableOf(shape.source);
	if (!tbl) return null;
	const range = sel.cells ?? { r0: 0, c0: 0, r1: 0, c1: 0 };
	return { shape, tbl, range };
}

export function hasTableSelection(ctx: CommandContext | null, sel: TableSelection): boolean {
	return tableTarget(ctx, sel) !== null;
}

function parts(ctx: CommandContext): string[] {
	return [ctx.slide.partPath, relsPathFor(ctx.slide.partPath)];
}

export function insertTableRow(
	ctx: CommandContext,
	sel: TableSelection,
	where: "above" | "below",
): boolean {
	const target = tableTarget(ctx, sel);
	if (!target) return false;
	const at = where === "above" ? target.range.r0 : target.range.r1 + 1;
	return ctx.editor.transact(
		where === "above" ? "Insert row above" : "Insert row below",
		parts(ctx),
		() => insertRow(target.tbl, at, target.range.r0),
	);
}

export function insertTableColumn(
	ctx: CommandContext,
	sel: TableSelection,
	where: "left" | "right",
): boolean {
	const target = tableTarget(ctx, sel);
	if (!target) return false;
	const at = where === "left" ? target.range.c0 : target.range.c1 + 1;
	return ctx.editor.transact(
		where === "left" ? "Insert column left" : "Insert column right",
		parts(ctx),
		() => insertColumn(target.tbl, at, target.range.c0),
	);
}

export function deleteTableRows(ctx: CommandContext, sel: TableSelection): boolean {
	const target = tableTarget(ctx, sel);
	if (!target) return false;
	return ctx.editor.transact("Delete rows", parts(ctx), () => {
		let changed = false;
		// Remove from the bottom so earlier indices stay valid.
		for (let r = target.range.r1; r >= target.range.r0; r--) {
			if (deleteRow(target.tbl, r)) changed = true;
		}
		return changed;
	});
}

export function deleteTableColumns(ctx: CommandContext, sel: TableSelection): boolean {
	const target = tableTarget(ctx, sel);
	if (!target) return false;
	return ctx.editor.transact("Delete columns", parts(ctx), () => {
		let changed = false;
		for (let c = target.range.c1; c >= target.range.c0; c--) {
			if (deleteColumn(target.tbl, c)) changed = true;
		}
		return changed;
	});
}

export function mergeTableCells(ctx: CommandContext, sel: TableSelection): boolean {
	const target = tableTarget(ctx, sel);
	if (!target) return false;
	return ctx.editor.transact("Merge cells", parts(ctx), () =>
		mergeCells(target.tbl, target.range),
	);
}

export function splitTableCells(ctx: CommandContext, sel: TableSelection): boolean {
	const target = tableTarget(ctx, sel);
	if (!target) return false;
	return ctx.editor.transact("Split cells", parts(ctx), () =>
		splitCells(target.tbl, target.range),
	);
}

export function setCellFill(
	ctx: CommandContext,
	sel: TableSelection,
	color: string | null,
): boolean {
	const target = tableTarget(ctx, sel);
	if (!target) return false;
	const doc = target.tbl.ownerDocument;

	return ctx.editor.transact(color === null ? "Clear cell fill" : "Cell fill", parts(ctx), () => {
		let changed = false;
		for (let r = target.range.r0; r <= target.range.r1; r++) {
			for (let c = target.range.c0; c <= target.range.c1; c++) {
				const tc = cellAt(target.tbl, r, c);
				if (!tc) continue;
				let props: Element | null = null;
				for (let n = tc.firstElementChild; n; n = n.nextElementSibling) {
					if (n.localName === "tcPr") props = n;
				}
				if (!props) {
					props = doc.createElementNS(A_NS, "a:tcPr");
					tc.appendChild(props);
				}
				for (const name of ["noFill", "solidFill", "gradFill", "blipFill", "pattFill"]) {
					setOrderedChild(props, name, TCPR_ORDER, null);
				}
				if (color === null) {
					setOrderedChild(props, "noFill", TCPR_ORDER, doc.createElementNS(A_NS, "a:noFill"));
				} else {
					const fill = doc.createElementNS(A_NS, "a:solidFill");
					const clr = doc.createElementNS(A_NS, "a:srgbClr");
					clr.setAttribute("val", color.replace(/^#/, "").toUpperCase());
					fill.appendChild(clr);
					setOrderedChild(props, "solidFill", TCPR_ORDER, fill);
				}
				changed = true;
			}
		}
		return changed;
	});
}

/** Row and column counts, for enabling and labelling ribbon controls. */
export function tableSize(
	ctx: CommandContext | null,
	sel: TableSelection,
): { rows: number; columns: number } | null {
	const target = tableTarget(ctx, sel);
	if (!target) return null;
	return { rows: tableRows(target.tbl).length, columns: gridColumns(target.tbl).length };
}
