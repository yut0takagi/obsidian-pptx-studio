import { Notice, setIcon, setTooltip } from "obsidian";
import type { CommandContext } from "../edit/commands";
import { renameShape, selectedShapes } from "../edit/commands";
import { setGeometry } from "../edit/formatCommands";
import { t } from "../i18n";
import type { Shape } from "../pptx/types";
import { EMU_PER_PX } from "../pptx/types";
import { attr, child, numAttr } from "../pptx/xml";
import { buildPaneHeader } from "./paneHeader";

export interface InspectorPaneOptions {
	getContext: () => CommandContext | null;
	run: (fn: (ctx: CommandContext) => unknown) => void;
	/** Start folded away, and report a fold back so it can be remembered. */
	collapsed?: boolean;
	onCollapsed?: (collapsed: boolean) => void;
}

/** The a:xfrm a shape's position is written into, wherever it lives. */
function xfrmOf(source: Element): Element | null {
	switch (source.localName) {
		case "grpSp":
			return child(child(source, "grpSpPr"), "xfrm");
		case "graphicFrame":
			// A graphic frame carries p:xfrm directly rather than inside spPr.
			return child(source, "xfrm");
		default:
			return child(child(source, "spPr"), "xfrm");
	}
}

/**
 * What the selected shape actually is, underneath the picture of it.
 *
 * The rendered slide answers "where is this box"; this answers "what does the
 * file say", which is the question you are asking when the two disagree — the
 * shape id to match against the XML, the raw EMU behind a rounded pixel
 * position, whether a frame is really there or inherited from the layout.
 *
 * The frame, the rotation and the name are editable here, because reading a
 * number off the file and then hunting for the control that sets it is the
 * long way round: type over the value. Everything a field writes goes through
 * the same commands the ribbon uses, so it undoes like any other edit. What is
 * structural — the shape id, the part, the element — stays read-only.
 */
export class InspectorPane {
	private readonly root: HTMLElement;
	private readonly bodyEl: HTMLElement;
	/** The XML on screen, for the copy button. */
	private xml = "";

	constructor(
		containerEl: HTMLElement,
		private readonly options: InspectorPaneOptions,
	) {
		this.root = containerEl.createDiv({ cls: "pptx-inspector-pane" });
		const head = buildPaneHeader(this.root, t("pane.properties"), {
			collapsed: options.collapsed ?? false,
			onToggle: (collapsed) => options.onCollapsed?.(collapsed),
		});
		const copy = head.createDiv({ cls: "pptx-pane-copy clickable-icon" });
		setIcon(copy, "copy");
		setTooltip(copy, t("pane.copyXml"));
		copy.addEventListener("click", (event) => {
			// The header itself folds the pane; the button inside it must not.
			event.stopPropagation();
			if (!this.xml) return;
			void navigator.clipboard.writeText(this.xml);
			new Notice(t("pane.copied"));
		});
		this.bodyEl = this.root.createDiv({ cls: "pptx-inspector-body" });
		this.refresh();
	}

	get element(): HTMLElement {
		return this.root;
	}

	refresh(): void {
		// Rebuilding the pane under a field someone is typing in would throw the
		// half-typed value away; the commit on blur brings it back into step.
		if (this.bodyEl.contains(document.activeElement)) return;
		this.bodyEl.empty();
		this.xml = "";
		const ctx = this.options.getContext();
		if (!ctx) return;
		const shapes = selectedShapes(ctx);
		if (shapes.length === 0) {
			this.bodyEl.createDiv({ cls: "pptx-pane-empty", text: t("pane.noSelection") });
			return;
		}
		if (shapes.length > 1) {
			this.buildSummary(shapes);
			return;
		}
		this.buildDetail(shapes[0], ctx);
	}

	/** Several shapes at once: the frames side by side, which is what gets compared. */
	private buildSummary(shapes: Shape[]): void {
		this.bodyEl.createDiv({
			cls: "pptx-pane-empty",
			text: t("prop.multiple", { n: shapes.length }),
		});
		const table = this.bodyEl.createDiv({ cls: "pptx-props" });
		for (const shape of shapes) {
			const f = shape.frame;
			this.addRow(
				table,
				`#${shape.id}`,
				`${Math.round(f.x)}, ${Math.round(f.y)} · ${Math.round(f.w)} × ${Math.round(f.h)}`,
			);
		}
	}

	private buildDetail(shape: Shape, ctx: CommandContext): void {
		const source = shape.source ?? null;
		const table = this.bodyEl.createDiv({ cls: "pptx-props" });
		const f = shape.frame;

		this.addRow(table, t("prop.id"), shape.id);
		this.addTextRow(table, t("prop.name"), shape.name, (value) => {
			this.options.run((c) => renameShape(c, shape.id, value));
		});
		this.addRow(table, t("prop.kind"), source ? `${shape.kind} · <${source.nodeName}>` : shape.kind);
		this.addRow(table, t("prop.part"), shape.sourcePart ?? ctx.slide.partPath);

		const ph = source ? child(child(child(source, "nvSpPr"), "nvPr"), "ph") : null;
		if (ph) {
			const type = attr(ph, "type") ?? "body";
			const index = attr(ph, "idx");
			this.addRow(table, t("prop.placeholder"), index ? `${type} (idx ${index})` : type);
		}

		const xfrm = source ? xfrmOf(source) : null;
		const off = child(xfrm, "off");
		const ext = child(xfrm, "ext");
		this.addLengthRow(table, "x", f.x, numAttr(off, "x"), (x) => ({ x }));
		this.addLengthRow(table, "y", f.y, numAttr(off, "y"), (y) => ({ y }));
		this.addLengthRow(table, "cx", f.w, numAttr(ext, "cx"), (w) => ({ w }));
		this.addLengthRow(table, "cy", f.h, numAttr(ext, "cy"), (h) => ({ h }));

		const rot = numAttr(xfrm, "rot");
		this.addFieldRow(
			table,
			t("prop.rotation"),
			round(f.rot),
			rot === null ? "" : `rot="${rot}"`,
			t("prop.degreesHint"),
			(text) => {
				const degrees = Number(text);
				if (!Number.isFinite(degrees)) return;
				this.options.run((c) => setGeometry(c, { rotation: degrees }, "Set rotation"));
			},
		);
		if (f.flipH || f.flipV) {
			this.addRow(table, t("prop.flip"), [f.flipH ? "H" : "", f.flipV ? "V" : ""].join(" ").trim());
		}
		if (shape.hidden) this.addRow(table, t("prop.hidden"), "true");
		if (!xfrm) {
			this.bodyEl.createDiv({ cls: "pptx-pane-note", text: t("prop.inherited") });
		}

		if (!source) return;
		this.xml = formatXml(source);
		const pre = this.bodyEl.createEl("pre", { cls: "pptx-xml" });
		pre.setText(
			this.xml.length > XML_LIMIT
				? `${this.xml.slice(0, XML_LIMIT)}\n…\n${t("prop.truncated")}`
				: this.xml,
		);
	}

	/**
	 * A length, editable in pixels but shown alongside the EMU actually stored —
	 * which is the number in the file, not a conversion of the rounded pixels, so
	 * that what the pane reports can be matched against the XML character for
	 * character. A unit suffix on the way in is honoured, `2743200emu` included,
	 * for setting a value read out of someone else's file.
	 */
	private addLengthRow(
		table: HTMLElement,
		label: string,
		px: number,
		emu: number | null,
		patch: (px: number) => { x?: number; y?: number; w?: number; h?: number },
	): void {
		const stored = emu ?? Math.round(px * EMU_PER_PX);
		this.addFieldRow(
			table,
			label,
			round(px),
			`${stored.toLocaleString("en-US")} EMU`,
			t("prop.lengthHint"),
			(text) => {
				const value = parseLength(text);
				if (value === null) return;
				this.options.run((c) => setGeometry(c, patch(value), "Set geometry"));
			},
		);
	}

	private addTextRow(
		table: HTMLElement,
		label: string,
		value: string,
		apply: (value: string) => void,
	): void {
		this.addFieldRow(table, label, value, "", "", (text) => {
			const trimmed = text.trim();
			if (trimmed && trimmed !== value) apply(trimmed);
		});
	}

	/**
	 * A row whose value can be typed over. The edit lands on Enter or on leaving
	 * the field, and Escape puts the old text back — the same bargain as the
	 * rename in the shape list.
	 */
	private addFieldRow(
		table: HTMLElement,
		label: string,
		value: string,
		note: string,
		hint: string,
		apply: (text: string) => void,
	): void {
		const row = table.createDiv({ cls: "pptx-prop" });
		row.createDiv({ cls: "pptx-prop-key", text: label });
		const field = row.createDiv({ cls: "pptx-prop-value is-editable" });
		const input = field.createEl("input", { cls: "pptx-prop-input" });
		input.type = "text";
		input.value = value;
		if (hint) setTooltip(input, hint);
		if (note) field.createSpan({ cls: "pptx-prop-note", text: note });

		let settled = false;
		input.addEventListener("blur", () => {
			if (settled) return;
			settled = true;
			if (input.value !== value) apply(input.value);
		});
		// The deck view listens for keys on an ancestor: unstopped, typing "j"
		// here would page the slide and the arrow keys would nudge the shape.
		input.addEventListener("keydown", (event) => {
			event.stopPropagation();
			if (event.key === "Enter") {
				event.preventDefault();
				input.blur();
			} else if (event.key === "Escape") {
				event.preventDefault();
				input.value = value;
				input.blur();
			}
		});
	}

	private addRow(table: HTMLElement, label: string, value: string): void {
		const row = table.createDiv({ cls: "pptx-prop" });
		row.createDiv({ cls: "pptx-prop-key", text: label });
		row.createDiv({ cls: "pptx-prop-value", text: value });
	}

	destroy(): void {
		this.root.detach();
	}
}

const XML_LIMIT = 20000;

const round = (value: number): string => (Math.round(value * 10) / 10).toString();

/** Pixels per unit, for the suffixes a length field accepts. */
const UNITS: Record<string, number> = {
	px: 1,
	emu: 1 / EMU_PER_PX,
	pt: 96 / 72,
	in: 96,
	cm: 96 / 2.54,
	mm: 96 / 25.4,
};

/** A typed length in pixels: a bare number is pixels, a suffix is honoured. */
function parseLength(text: string): number | null {
	const match = /^\s*(-?\d+(?:\.\d+)?)\s*([a-z]*)\s*$/i.exec(text);
	if (!match) return null;
	const value = Number(match[1]);
	const unit = match[2].toLowerCase();
	if (!Number.isFinite(value)) return null;
	if (!unit) return value;
	const factor = UNITS[unit];
	return factor === undefined ? null : value * factor;
}

/** Serialised XML, broken onto one element per line and indented to match. */
function formatXml(el: Element): string {
	const raw = new XMLSerializer().serializeToString(el);
	let depth = 0;
	return raw
		.replace(/>\s*</g, ">\n<")
		.split("\n")
		.map((line) => {
			const closing = /^<\//.test(line);
			if (closing) depth = Math.max(0, depth - 1);
			const out = "  ".repeat(depth) + line;
			// An opening tag with nothing else on the line indents what follows;
			// `<a:t>text</a:t>` and `<a:off …/>` are complete in themselves.
			if (!closing && /^<[^/][^>]*[^/]>$/.test(line)) depth += 1;
			return out;
		})
		.join("\n");
}
