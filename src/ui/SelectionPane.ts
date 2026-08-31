import { setIcon, setTooltip } from "obsidian";
import type { CommandContext } from "../edit/commands";
import { renameShape, reorderSelection, setShapeHidden, shapeListing } from "../edit/commands";
import type { Selection } from "../edit/Selection";
import { t } from "../i18n";
import { buildPaneHeader } from "./paneHeader";
import type { Shape } from "../pptx/types";

export interface SelectionPaneOptions {
	getContext: () => CommandContext | null;
	selection: Selection;
	run: (fn: (ctx: CommandContext) => unknown) => void;
	/** Start folded away, and report a fold back so it can be remembered. */
	collapsed?: boolean;
	onCollapsed?: (collapsed: boolean) => void;
}

/**
 * A list of everything on the slide, in stacking order.
 *
 * It exists for the shapes direct manipulation cannot reach: something hidden,
 * something exactly behind something else, something with no visible pixels.
 * Top-most first, matching how the slide is read rather than how it is stored.
 */
export class SelectionPane {
	private readonly root: HTMLElement;
	private readonly listEl: HTMLElement;

	constructor(
		containerEl: HTMLElement,
		private readonly options: SelectionPaneOptions,
	) {
		this.root = containerEl.createDiv({ cls: "pptx-selection-pane" });
		buildPaneHeader(this.root, t("pane.selection"), {
			collapsed: options.collapsed ?? false,
			onToggle: (collapsed) => options.onCollapsed?.(collapsed),
		});
		this.listEl = this.root.createDiv({ cls: "pptx-pane-list" });
		this.refresh();
	}

	get element(): HTMLElement {
		return this.root;
	}

	refresh(): void {
		const ctx = this.options.getContext();
		this.listEl.empty();
		if (!ctx) return;
		const shapes = shapeListing(ctx);
		if (shapes.length === 0) {
			this.listEl.createDiv({ cls: "pptx-pane-empty", text: t("pane.empty") });
			return;
		}
		for (const shape of shapes) this.buildRow(shape);
	}

	private buildRow(shape: Shape): void {
		const row = this.listEl.createDiv({ cls: "pptx-pane-row" });
		row.toggleClass("is-selected", this.options.selection.has(shape.id));
		row.toggleClass("is-hidden", shape.hidden);

		const eye = row.createDiv({ cls: "pptx-pane-eye clickable-icon" });
		setIcon(eye, shape.hidden ? "eye-off" : "eye");
		setTooltip(eye, shape.hidden ? t("pane.show") : t("pane.hide"));
		eye.addEventListener("click", (event) => {
			event.stopPropagation();
			this.options.run((ctx) => setShapeHidden(ctx, shape.id, !shape.hidden));
		});

		const label = row.createDiv({ cls: "pptx-pane-name", text: shape.name || describe(shape) });
		label.addEventListener("dblclick", (event) => {
			event.stopPropagation();
			this.startRename(label, shape);
		});

		row.addEventListener("click", (event) => {
			const additive = event.shiftKey || event.metaKey || event.ctrlKey;
			const ctx = this.options.getContext();
			if (!ctx) return;
			const slideIndex = ctx.slide.index - 1;
			if (additive) this.options.selection.toggle(slideIndex, shape.id);
			else this.options.selection.set(slideIndex, [shape.id]);
		});

		const up = row.createDiv({ cls: "pptx-pane-move clickable-icon" });
		setIcon(up, "chevron-up");
		setTooltip(up, t("cmd.bringForward"));
		up.addEventListener("click", (event) => {
			event.stopPropagation();
			this.reorder(shape, "forward");
		});

		const down = row.createDiv({ cls: "pptx-pane-move clickable-icon" });
		setIcon(down, "chevron-down");
		setTooltip(down, t("cmd.sendBackward"));
		down.addEventListener("click", (event) => {
			event.stopPropagation();
			this.reorder(shape, "backward");
		});
	}

	/** Reordering acts on the clicked row, whatever happened to be selected. */
	private reorder(shape: Shape, mode: "forward" | "backward"): void {
		const ctx = this.options.getContext();
		if (!ctx) return;
		this.options.selection.set(ctx.slide.index - 1, [shape.id]);
		this.options.run((c) => reorderSelection(c, mode));
	}

	private startRename(label: HTMLElement, shape: Shape): void {
		const input = createEl("input");
		input.type = "text";
		input.value = shape.name;
		input.addClass("pptx-pane-rename");
		label.replaceWith(input);
		input.focus();
		input.select();

		const commit = (): void => {
			const value = input.value.trim();
			input.remove();
			if (value && value !== shape.name) {
				this.options.run((ctx) => renameShape(ctx, shape.id, value));
			} else {
				this.refresh();
			}
		};
		input.addEventListener("blur", commit);
		input.addEventListener("keydown", (event) => {
			event.stopPropagation();
			if (event.key === "Enter") commit();
			if (event.key === "Escape") {
				input.remove();
				this.refresh();
			}
		});
	}

	destroy(): void {
		this.root.detach();
	}
}

/** A fallback name for a shape that never got one. */
function describe(shape: Shape): string {
	switch (shape.kind) {
		case "image":
			return t("cmd.picture");
		case "table":
			return t("cmd.table");
		case "chart":
			return shape.chart.title || t("pane.chart");
		case "group":
			return t("cmd.group");
		case "line":
			return t("cmd.line");
		default:
			return t("cmd.textBox");
	}
}
