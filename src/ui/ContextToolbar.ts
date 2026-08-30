import { Menu, setIcon, setTooltip } from "obsidian";
import type { CommandContext } from "../edit/commands";
import {
	alignSelection,
	deleteSelection,
	duplicateSelection,
	groupSelection,
	reorderSelection,
	selectedShapes,
	ungroupSelection,
} from "../edit/commands";
import {
	applyTextFormat,
	fillState,
	flipSelection,
	outlineState,
	rotateBy,
	setFill,
	setOutline,
	textState,
} from "../edit/formatCommands";
import type { Selection } from "../edit/Selection";
import {
	type TableSelection,
	deleteTableColumns,
	deleteTableRows,
	insertTableColumn,
	insertTableRow,
	mergeTableCells,
	splitTableCells,
} from "../edit/tableCommands";
import { t } from "../i18n";
import { openColorPopup } from "./Ribbon";

export interface ContextToolbarOptions {
	getContext: () => CommandContext | null;
	selection: Selection;
	tableSelection: TableSelection;
	run: (fn: (ctx: CommandContext) => unknown) => void;
	getSlideEl: () => HTMLElement | null;
	/** The stage's client rect, so the bar never floats outside the canvas. */
	getViewportEl: () => HTMLElement | null;
	getScale: () => number;
	isEditing: () => boolean;
	canCrop: () => boolean;
	cropActive: () => boolean;
	toggleCrop: () => void;
}

/**
 * A small toolbar that follows the selection.
 *
 * The ribbon is a long way from the slide in a pane this size, and the actions
 * people reach for constantly — fill, delete, duplicate, bold — are worth having
 * under the pointer. Everything here also exists in the ribbon; this is a
 * shortcut, not a second home for any command.
 */
export class ContextToolbar {
	private readonly root: HTMLElement;

	constructor(private readonly options: ContextToolbarOptions) {
		this.root = document.body.createDiv({ cls: "pptx-context-toolbar is-hidden" });
		// Pressing a button must not pull focus off the canvas.
		this.root.addEventListener("mousedown", (event) => event.preventDefault());
	}

	refresh(): void {
		const ctx = this.options.getContext();
		const shapes = ctx ? selectedShapes(ctx) : [];
		if (!ctx || shapes.length === 0 || this.options.isEditing()) {
			this.root.addClass("is-hidden");
			return;
		}
		this.build(ctx, shapes.length);
		this.position();
	}

	private build(ctx: CommandContext, count: number): void {
		this.root.empty();

		const shapes = selectedShapes(ctx);
		const hasText = textState(ctx) !== null;
		const hasGroup = shapes.some((s) => s.kind === "group");
		const isTable = shapes.length === 1 && shapes[0].kind === "table";

		if (hasText) {
			this.toggleButton("bold", t("cmd.bold"), () => textState(ctx)?.bold ?? false, () =>
				this.options.run((c) =>
					applyTextFormat(c, { bold: !(textState(c)?.bold ?? false) }, t("cmd.bold")),
				),
			);
			this.toggleButton("italic", t("cmd.italic"), () => textState(ctx)?.italic ?? false, () =>
				this.options.run((c) =>
					applyTextFormat(c, { italic: !(textState(c)?.italic ?? false) }, t("cmd.italic")),
				),
			);
			this.separator();
		}

		this.colourButton("paint-bucket", t("cmd.shapeFill"), () => fillState(this.options.getContext()), (colour) =>
			this.options.run((c) => setFill(c, colour)),
		);
		this.colourButton("square", t("cmd.shapeOutline"), () => outlineState(this.options.getContext()), (colour) =>
			this.options.run((c) => setOutline(c, { color: colour }, t("cmd.shapeOutline"))),
		);
		this.separator();

		if (this.options.canCrop() || this.options.cropActive()) {
			this.toggleButton("crop", t("cmd.crop"), this.options.cropActive, this.options.toggleCrop);
		}

		if (isTable) {
			this.menuButton("table", t("tab.table"), (menu) => {
				const sel = this.options.tableSelection;
				menu.addItem((item) =>
					item
						.setTitle(t("cmd.rowAboveTooltip"))
						.onClick(() => this.options.run((c) => insertTableRow(c, sel, "above"))),
				);
				menu.addItem((item) =>
					item
						.setTitle(t("cmd.rowBelowTooltip"))
						.onClick(() => this.options.run((c) => insertTableRow(c, sel, "below"))),
				);
				menu.addItem((item) =>
					item
						.setTitle(t("cmd.columnLeftTooltip"))
						.onClick(() => this.options.run((c) => insertTableColumn(c, sel, "left"))),
				);
				menu.addItem((item) =>
					item
						.setTitle(t("cmd.columnRightTooltip"))
						.onClick(() => this.options.run((c) => insertTableColumn(c, sel, "right"))),
				);
				menu.addSeparator();
				menu.addItem((item) =>
					item
						.setTitle(t("cmd.mergeTooltip"))
						.onClick(() => this.options.run((c) => mergeTableCells(c, sel))),
				);
				menu.addItem((item) =>
					item
						.setTitle(t("cmd.splitTooltip"))
						.onClick(() => this.options.run((c) => splitTableCells(c, sel))),
				);
				menu.addSeparator();
				menu.addItem((item) =>
					item
						.setTitle(t("cmd.deleteRows"))
						.onClick(() => this.options.run((c) => deleteTableRows(c, sel))),
				);
				menu.addItem((item) =>
					item
						.setTitle(t("cmd.deleteColumns"))
						.onClick(() => this.options.run((c) => deleteTableColumns(c, sel))),
				);
			});
		}

		this.menuButton("layers", t("cmd.arrangeTooltip"), (menu) => {
			for (const option of [
				{ label: t("cmd.bringToFront"), mode: "front" as const, icon: "bring-to-front" },
				{ label: t("cmd.bringForward"), mode: "forward" as const, icon: "arrow-up" },
				{ label: t("cmd.sendBackward"), mode: "backward" as const, icon: "arrow-down" },
				{ label: t("cmd.sendToBack"), mode: "back" as const, icon: "send-to-back" },
			]) {
				menu.addItem((item) =>
					item
						.setTitle(option.label)
						.setIcon(option.icon)
						.onClick(() => this.options.run((c) => reorderSelection(c, option.mode))),
				);
			}
			menu.addSeparator();
			for (const option of [
				{ label: t("cmd.alignLeft"), mode: "left" as const },
				{ label: t("cmd.alignCentre"), mode: "centerH" as const },
				{ label: t("cmd.alignRight"), mode: "right" as const },
				{ label: t("cmd.alignTop"), mode: "top" as const },
				{ label: t("cmd.alignMiddle"), mode: "middle" as const },
				{ label: t("cmd.alignBottom"), mode: "bottom" as const },
			]) {
				menu.addItem((item) =>
					item
						.setTitle(option.label)
						.onClick(() => this.options.run((c) => alignSelection(c, option.mode))),
				);
			}
			menu.addSeparator();
			menu.addItem((item) =>
				item
					.setTitle(t("cmd.rotateRight"))
					.setIcon("rotate-cw")
					.onClick(() => this.options.run((c) => rotateBy(c, 90))),
			);
			menu.addItem((item) =>
				item
					.setTitle(t("cmd.flipH"))
					.setIcon("flip-horizontal")
					.onClick(() => this.options.run((c) => flipSelection(c, "h"))),
			);
			if (count > 1) {
				menu.addSeparator();
				menu.addItem((item) =>
					item
						.setTitle(t("cmd.group"))
						.setIcon("group")
						.onClick(() => this.options.run(groupSelection)),
				);
			}
			if (hasGroup) {
				menu.addItem((item) =>
					item
						.setTitle(t("cmd.ungroup"))
						.setIcon("ungroup")
						.onClick(() => this.options.run(ungroupSelection)),
				);
			}
		});

		this.separator();
		this.button("copy-plus", t("cmd.duplicate"), () => this.options.run(duplicateSelection));
		this.button("trash", t("cmd.delete"), () => this.options.run(deleteSelection));
	}

	// ------------------------------------------------------------ controls

	private button(icon: string, tooltip: string, onClick: () => void): HTMLElement {
		const el = this.root.createEl("button", { cls: "pptx-ctx-btn" });
		setIcon(el, icon);
		setTooltip(el, tooltip);
		el.addEventListener("click", (event) => {
			event.preventDefault();
			onClick();
		});
		return el;
	}

	private toggleButton(
		icon: string,
		tooltip: string,
		isActive: () => boolean,
		onClick: () => void,
	): void {
		const el = this.button(icon, tooltip, onClick);
		el.toggleClass("is-active", isActive());
	}

	private menuButton(icon: string, tooltip: string, build: (menu: Menu) => void): void {
		const el = this.button(icon, tooltip, () => {
			const menu = new Menu();
			build(menu);
			const rect = el.getBoundingClientRect();
			menu.showAtPosition({ x: rect.left, y: rect.bottom + 4 });
		});
		el.addClass("has-menu");
	}

	private colourButton(
		icon: string,
		tooltip: string,
		value: () => string | null,
		onChange: (colour: string | null) => void,
	): void {
		const el = this.button(icon, tooltip, () => {
			openColorPopup(el, {
				kind: "color",
				icon,
				tooltip,
				allowNone: true,
				value,
				onChange,
			});
		});
		const swatch = el.createDiv({ cls: "pptx-ctx-swatch" });
		swatch.style.background = value() ?? "transparent";
		swatch.toggleClass("is-none", value() === null);
	}

	private separator(): void {
		this.root.createDiv({ cls: "pptx-ctx-sep" });
	}

	// ----------------------------------------------------------- placement

	/** Sit just above the selection, flipping below it when there is no room. */
	private position(): void {
		const slideEl = this.options.getSlideEl();
		const viewport = this.options.getViewportEl();
		const ctx = this.options.getContext();
		// The bar lives on document.body, so it has to check for itself whether the
		// view it belongs to is still on screen — a background tab must not leave
		// a toolbar floating over whatever the user switched to.
		if (!slideEl?.isConnected || !viewport?.offsetParent || !ctx) {
			this.root.addClass("is-hidden");
			return;
		}
		const shapes = selectedShapes(ctx);
		if (shapes.length === 0) {
			this.root.addClass("is-hidden");
			return;
		}

		const scale = Math.max(this.options.getScale(), 0.05);
		const slideRect = slideEl.getBoundingClientRect();
		const left = Math.min(...shapes.map((s) => s.frame.x));
		const right = Math.max(...shapes.map((s) => s.frame.x + s.frame.w));
		const top = Math.min(...shapes.map((s) => s.frame.y));
		const bottom = Math.max(...shapes.map((s) => s.frame.y + s.frame.h));

		const centre = slideRect.left + ((left + right) / 2) * scale;
		const above = slideRect.top + top * scale;
		const below = slideRect.top + bottom * scale;

		const bar = this.root.getBoundingClientRect();
		const view = viewport.getBoundingClientRect();
		const gap = 10;

		let y = above - bar.height - gap;
		if (y < view.top + 4) y = Math.min(below + gap, view.bottom - bar.height - 4);
		y = Math.max(view.top + 4, Math.min(y, view.bottom - bar.height - 4));

		const x = Math.max(
			view.left + 4,
			Math.min(centre - bar.width / 2, view.right - bar.width - 4),
		);

		// A selection scrolled out of the stage should not leave the bar hanging.
		const visible = below > view.top - 40 && above < view.bottom + 40;
		this.root.toggleClass("is-hidden", !visible);
		this.root.style.left = `${Math.round(x)}px`;
		this.root.style.top = `${Math.round(y)}px`;
	}

	destroy(): void {
		this.root.detach();
	}
}
