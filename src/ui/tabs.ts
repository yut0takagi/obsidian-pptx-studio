import type { Menu } from "obsidian";
import type { CommandContext } from "../edit/commands";
import {
	alignSelection,
	copySelection,
	cutSelection,
	deleteSelection,
	distributeSelection,
	duplicateSelection,
	groupSelection,
	hasClipboard,
	pasteClipboard,
	reorderSelection,
	selectedShapes,
	ungroupSelection,
} from "../edit/commands";
import {
	applyParagraphFormat,
	applyTextFormat,
	fillState,
	outlineState,
	setFill,
	setOutline,
	setVerticalAnchor,
	textState,
} from "../edit/formatCommands";
import { insertAutoShape, insertLine, insertTextBox } from "../edit/insertCommands";
import {
	canDeleteSlide,
	deleteCurrentSlide,
	duplicateCurrentSlide,
	moveCurrentSlide,
	newSlide,
} from "../edit/slideCommands";
import type { RibbonItem, RibbonTab } from "./Ribbon";

export interface RibbonHost {
	ctx: () => CommandContext | null;
	/** Run a command and refresh the UI around it. */
	run: (fn: (ctx: CommandContext) => unknown) => void;
	/** Run a slide-level command whose result is the slide to show afterwards. */
	runSlide: (fn: (ctx: CommandContext) => number) => void;
	canEdit: () => boolean;
	zoomIn: () => void;
	zoomOut: () => void;
	zoomToFit: () => void;
	toggleNotes: () => void;
	notesShown: () => boolean;
	toggleThumbnails: () => void;
	save: () => void;
	isDirty: () => boolean;
	undo: () => void;
	redo: () => void;
	canUndo: () => boolean;
	canRedo: () => boolean;
	selectAll: () => void;
	pickImage: () => void;
	pickTable: () => void;
	pickLayout: () => void;
	exportPng: () => void;
	extractMarkdown: () => void;
	openExternally: () => void;
}

/** Presets offered by the shapes menu, in the order PowerPoint groups them. */
export const SHAPE_PRESETS: { preset: string; label: string }[] = [
	{ preset: "rect", label: "Rectangle" },
	{ preset: "roundRect", label: "Rounded rectangle" },
	{ preset: "ellipse", label: "Ellipse" },
	{ preset: "triangle", label: "Triangle" },
	{ preset: "diamond", label: "Diamond" },
	{ preset: "pentagon", label: "Pentagon" },
	{ preset: "hexagon", label: "Hexagon" },
	{ preset: "star5", label: "Star" },
	{ preset: "rightArrow", label: "Arrow" },
	{ preset: "chevron", label: "Chevron" },
	{ preset: "flowChartDecision", label: "Decision" },
	{ preset: "wedgeRectCallout", label: "Callout" },
];

const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 54, 60, 66, 72, 88];

const FONTS = [
	"Helvetica Neue",
	"Arial",
	"Calibri",
	"Georgia",
	"Times New Roman",
	"Courier New",
	"Verdana",
	"Hiragino Sans",
	"Yu Gothic",
	"Meiryo",
	"Noto Sans JP",
];

export function buildTabs(host: RibbonHost): RibbonTab[] {
	const hasSelection = () => (host.ctx() ? selectedShapes(host.ctx()!).length > 0 : false);
	const selectionCount = () => (host.ctx() ? selectedShapes(host.ctx()!).length : 0);
	const hasGroup = () =>
		host.ctx() ? selectedShapes(host.ctx()!).some((s) => s.kind === "group") : false;
	const state = () => textState(host.ctx());

	const textButton = (
		icon: string,
		tooltip: string,
		key: "bold" | "italic" | "underline" | "strike",
	): RibbonItem => ({
		kind: "button",
		icon,
		tooltip,
		isEnabled: () => host.canEdit() && hasSelection(),
		isActive: () => state()?.[key] ?? false,
		onClick: () =>
			host.run((ctx) => applyTextFormat(ctx, { [key]: !(state()?.[key] ?? false) }, tooltip)),
	});

	const alignButton = (icon: string, tooltip: string, value: "l" | "ctr" | "r" | "just"): RibbonItem => ({
		kind: "button",
		icon,
		tooltip,
		isEnabled: () => host.canEdit() && hasSelection(),
		isActive: () => {
			const current = state()?.align;
			const map: Record<string, string> = { left: "l", center: "ctr", right: "r", justify: "just" };
			return current !== undefined && current !== null && map[current] === value;
		},
		onClick: () => host.run((ctx) => applyParagraphFormat(ctx, { align: value }, tooltip)),
	});

	const fontGroup: RibbonItem[] = [
		{
			kind: "select",
			tooltip: "Font",
			width: "9.5em",
			options: () => FONTS.map((f) => ({ value: f, label: f })),
			value: () => state()?.font ?? "",
			isEnabled: () => host.canEdit() && hasSelection(),
			onChange: (value) => host.run((ctx) => applyTextFormat(ctx, { font: value }, "Font")),
		},
		{
			kind: "select",
			tooltip: "Font size",
			width: "4.5em",
			options: () => FONT_SIZES.map((s) => ({ value: String(s), label: String(s) })),
			value: () => {
				const size = state()?.size;
				return size === null || size === undefined ? "" : String(Math.round(size));
			},
			isEnabled: () => host.canEdit() && hasSelection(),
			onChange: (value) =>
				host.run((ctx) => applyTextFormat(ctx, { size: Number(value) }, "Font size")),
		},
		textButton("bold", "Bold", "bold"),
		textButton("italic", "Italic", "italic"),
		textButton("underline", "Underline", "underline"),
		textButton("strikethrough", "Strikethrough", "strike"),
		{
			kind: "color",
			icon: "baseline",
			tooltip: "Text colour",
			value: () => state()?.color ?? null,
			isEnabled: () => host.canEdit() && hasSelection(),
			onChange: (color) => host.run((ctx) => applyTextFormat(ctx, { color }, "Text colour")),
		},
	];

	const paragraphGroup: RibbonItem[] = [
		{
			kind: "button",
			icon: "list",
			tooltip: "Bulleted list",
			isEnabled: () => host.canEdit() && hasSelection(),
			isActive: () => state()?.bulleted ?? false,
			onClick: () =>
				host.run((ctx) =>
					applyParagraphFormat(
						ctx,
						{ bullet: state()?.bulleted ? "none" : "char" },
						"Bulleted list",
					),
				),
		},
		{
			kind: "button",
			icon: "list-ordered",
			tooltip: "Numbered list",
			isEnabled: () => host.canEdit() && hasSelection(),
			onClick: () => host.run((ctx) => applyParagraphFormat(ctx, { bullet: "number" }, "Numbered list")),
		},
		{
			kind: "button",
			icon: "indent-decrease",
			tooltip: "Decrease list level",
			isEnabled: () => host.canEdit() && hasSelection(),
			onClick: () => host.run((ctx) => applyParagraphFormat(ctx, { levelDelta: -1 }, "Outdent")),
		},
		{
			kind: "button",
			icon: "indent-increase",
			tooltip: "Increase list level",
			isEnabled: () => host.canEdit() && hasSelection(),
			onClick: () => host.run((ctx) => applyParagraphFormat(ctx, { levelDelta: 1 }, "Indent")),
		},
		{ kind: "separator" },
		alignButton("align-left", "Align left", "l"),
		alignButton("align-center", "Centre", "ctr"),
		alignButton("align-right", "Align right", "r"),
		alignButton("align-justify", "Justify", "just"),
		{
			kind: "menu",
			icon: "chevrons-up-down",
			tooltip: "Line spacing and vertical alignment",
			isEnabled: () => host.canEdit() && hasSelection(),
			build: (menu: Menu) => {
				for (const value of [1, 1.15, 1.5, 2]) {
					menu.addItem((item) =>
						item
							.setTitle(`Line spacing ${value}`)
							.onClick(() =>
								host.run((ctx) => applyParagraphFormat(ctx, { lineSpacing: value }, "Line spacing")),
							),
					);
				}
				menu.addSeparator();
				const anchors: { label: string; value: "t" | "ctr" | "b" }[] = [
					{ label: "Align text top", value: "t" },
					{ label: "Align text middle", value: "ctr" },
					{ label: "Align text bottom", value: "b" },
				];
				for (const anchor of anchors) {
					menu.addItem((item) =>
						item
							.setTitle(anchor.label)
							.onClick(() => host.run((ctx) => setVerticalAnchor(ctx, anchor.value))),
					);
				}
			},
		},
	];

	const arrangeItems: RibbonItem[] = [
		{
			kind: "menu",
			icon: "layers",
			label: "Arrange",
			tooltip: "Bring forward, send backward",
			isEnabled: () => host.canEdit() && hasSelection(),
			build: (menu: Menu) => {
				const options = [
					{ label: "Bring to front", mode: "front" as const, icon: "bring-to-front" },
					{ label: "Bring forward", mode: "forward" as const, icon: "arrow-up" },
					{ label: "Send backward", mode: "backward" as const, icon: "arrow-down" },
					{ label: "Send to back", mode: "back" as const, icon: "send-to-back" },
				];
				for (const option of options) {
					menu.addItem((item) =>
						item
							.setTitle(option.label)
							.setIcon(option.icon)
							.onClick(() => host.run((ctx) => reorderSelection(ctx, option.mode))),
					);
				}
			},
		},
		{
			kind: "menu",
			icon: "align-horizontal-justify-center",
			label: "Align",
			tooltip: "Align and distribute",
			isEnabled: () => host.canEdit() && hasSelection(),
			build: (menu: Menu) => {
				const options = [
					{ label: "Align left", mode: "left" as const },
					{ label: "Align centre", mode: "centerH" as const },
					{ label: "Align right", mode: "right" as const },
					{ label: "Align top", mode: "top" as const },
					{ label: "Align middle", mode: "middle" as const },
					{ label: "Align bottom", mode: "bottom" as const },
				];
				for (const option of options) {
					menu.addItem((item) =>
						item.setTitle(option.label).onClick(() => host.run((ctx) => alignSelection(ctx, option.mode))),
					);
				}
				menu.addSeparator();
				menu.addItem((item) =>
					item
						.setTitle("Distribute horizontally")
						.setDisabled(selectionCount() < 3)
						.onClick(() => host.run((ctx) => distributeSelection(ctx, "h"))),
				);
				menu.addItem((item) =>
					item
						.setTitle("Distribute vertically")
						.setDisabled(selectionCount() < 3)
						.onClick(() => host.run((ctx) => distributeSelection(ctx, "v"))),
				);
			},
		},
		{
			kind: "button",
			icon: "group",
			tooltip: "Group",
			isEnabled: () => host.canEdit() && selectionCount() > 1,
			onClick: () => host.run(groupSelection),
		},
		{
			kind: "button",
			icon: "ungroup",
			tooltip: "Ungroup",
			isEnabled: () => host.canEdit() && hasGroup(),
			onClick: () => host.run(ungroupSelection),
		},
	];

	const shapeStyleItems: RibbonItem[] = [
		{
			kind: "color",
			icon: "paint-bucket",
			tooltip: "Shape fill",
			allowNone: true,
			value: () => fillState(host.ctx()),
			isEnabled: () => host.canEdit() && hasSelection(),
			onChange: (color) => host.run((ctx) => setFill(ctx, color)),
		},
		{
			kind: "color",
			icon: "square",
			tooltip: "Shape outline",
			allowNone: true,
			value: () => outlineState(host.ctx()),
			isEnabled: () => host.canEdit() && hasSelection(),
			onChange: (color) => host.run((ctx) => setOutline(ctx, { color }, "Outline colour")),
		},
		{
			kind: "menu",
			icon: "pencil-ruler",
			tooltip: "Outline weight and style",
			isEnabled: () => host.canEdit() && hasSelection(),
			build: (menu: Menu) => {
				for (const width of [0.75, 1, 1.5, 2.25, 3, 4.5, 6]) {
					menu.addItem((item) =>
						item
							.setTitle(`${width} pt`)
							.onClick(() => host.run((ctx) => setOutline(ctx, { width }, "Outline weight"))),
					);
				}
				menu.addSeparator();
				const dashes = [
					{ label: "Solid", value: "solid" },
					{ label: "Dashed", value: "dash" },
					{ label: "Dotted", value: "sysDot" },
				];
				for (const dash of dashes) {
					menu.addItem((item) =>
						item
							.setTitle(dash.label)
							.onClick(() => host.run((ctx) => setOutline(ctx, { dash: dash.value }, "Outline style"))),
					);
				}
			},
		},
	];

	const shapesMenu: RibbonItem = {
		kind: "menu",
		icon: "shapes",
		label: "Shapes",
		tooltip: "Insert a shape",
		isEnabled: host.canEdit,
		build: (menu: Menu) => {
			for (const shape of SHAPE_PRESETS) {
				menu.addItem((item) =>
					item
						.setTitle(shape.label)
						.onClick(() => host.run((ctx) => insertAutoShape(ctx, shape.preset, shape.label))),
				);
			}
			menu.addSeparator();
			menu.addItem((item) =>
				item.setTitle("Line").onClick(() => host.run(insertLine)),
			);
		},
	};

	const clipboardItems: RibbonItem[] = [
		{
			kind: "button",
			icon: "clipboard-paste",
			tooltip: "Paste (Cmd/Ctrl+V)",
			isEnabled: () => host.canEdit() && hasClipboard(),
			onClick: () => host.run((ctx) => pasteClipboard(ctx)),
		},
		{
			kind: "button",
			icon: "scissors",
			tooltip: "Cut (Cmd/Ctrl+X)",
			isEnabled: () => host.canEdit() && hasSelection(),
			onClick: () => host.run(cutSelection),
		},
		{
			kind: "button",
			icon: "copy",
			tooltip: "Copy (Cmd/Ctrl+C)",
			isEnabled: () => host.canEdit() && hasSelection(),
			onClick: () => host.run(copySelection),
		},
		{
			kind: "button",
			icon: "copy-plus",
			tooltip: "Duplicate (Cmd/Ctrl+D)",
			isEnabled: () => host.canEdit() && hasSelection(),
			onClick: () => host.run(duplicateSelection),
		},
	];

	const slideItems: RibbonItem[] = [
		{
			kind: "button",
			icon: "file-plus",
			label: "New",
			tooltip: "New slide",
			isEnabled: host.canEdit,
			onClick: () => host.runSlide((ctx) => newSlide(ctx)),
		},
		{
			kind: "button",
			icon: "layout-template",
			tooltip: "New slide from a layout",
			isEnabled: host.canEdit,
			onClick: () => host.pickLayout(),
		},
		{
			kind: "button",
			icon: "files",
			tooltip: "Duplicate slide",
			isEnabled: host.canEdit,
			onClick: () => host.runSlide(duplicateCurrentSlide),
		},
		{
			kind: "button",
			icon: "trash-2",
			tooltip: "Delete slide",
			isEnabled: () => {
				const ctx = host.ctx();
				return host.canEdit() && ctx !== null && canDeleteSlide(ctx);
			},
			onClick: () => host.run(deleteCurrentSlide),
		},
		{
			kind: "button",
			icon: "arrow-up",
			tooltip: "Move slide up",
			isEnabled: host.canEdit,
			onClick: () => host.runSlide((ctx) => moveCurrentSlide(ctx, -1)),
		},
		{
			kind: "button",
			icon: "arrow-down",
			tooltip: "Move slide down",
			isEnabled: host.canEdit,
			onClick: () => host.runSlide((ctx) => moveCurrentSlide(ctx, 1)),
		},
	];

	return [
		{
			id: "home",
			title: "Home",
			groups: [
				{
					title: "Undo",
					items: [
						{
							kind: "button",
							icon: "undo-2",
							tooltip: "Undo (Cmd/Ctrl+Z)",
							isEnabled: host.canUndo,
							onClick: host.undo,
						},
						{
							kind: "button",
							icon: "redo-2",
							tooltip: "Redo (Cmd/Ctrl+Shift+Z)",
							isEnabled: host.canRedo,
							onClick: host.redo,
						},
						{
							kind: "button",
							icon: "save",
							tooltip: "Save (Cmd/Ctrl+S)",
							isEnabled: host.isDirty,
							onClick: host.save,
						},
					],
				},
				{ title: "Clipboard", items: clipboardItems },
				{ title: "Slides", items: slideItems.slice(0, 4) },
				{ title: "Font", items: fontGroup },
				{ title: "Paragraph", items: paragraphGroup },
				{ title: "Drawing", items: [shapesMenu, ...shapeStyleItems, ...arrangeItems] },
				{
					title: "Editing",
					items: [
						{
							kind: "button",
							icon: "mouse-pointer-2",
							tooltip: "Select all (Cmd/Ctrl+A)",
							isEnabled: host.canEdit,
							onClick: host.selectAll,
						},
						{
							kind: "button",
							icon: "trash",
							tooltip: "Delete (Del)",
							isEnabled: () => host.canEdit() && hasSelection(),
							onClick: () => host.run(deleteSelection),
						},
					],
				},
			],
		},
		{
			id: "insert",
			title: "Insert",
			groups: [
				{ title: "Slides", items: slideItems },
				{
					title: "Objects",
					items: [
						{
							kind: "button",
							icon: "type",
							label: "Text box",
							tooltip: "Insert a text box",
							isEnabled: host.canEdit,
							onClick: () => host.run(insertTextBox),
						},
						shapesMenu,
						{
							kind: "button",
							icon: "image",
							label: "Picture",
							tooltip: "Insert a picture from the vault",
							isEnabled: host.canEdit,
							onClick: () => host.pickImage(),
						},
						{
							kind: "button",
							icon: "table",
							label: "Table",
							tooltip: "Insert a table",
							isEnabled: host.canEdit,
							onClick: () => host.pickTable(),
						},
					],
				},
			],
		},
		{
			id: "format",
			title: "Format",
			groups: [
				{ title: "Shape styles", items: shapeStyleItems },
				{ title: "Arrange", items: arrangeItems },
				{
					title: "Quick shapes",
					items: [
						...SHAPE_PRESETS.slice(0, 6).map(
							(shape): RibbonItem => ({
								kind: "button",
								label: shape.label,
								tooltip: `Insert a ${shape.label.toLowerCase()}`,
								isEnabled: host.canEdit,
								onClick: () => host.run((ctx) => insertAutoShape(ctx, shape.preset, shape.label)),
							}),
						),
					],
				},
			],
		},
		{
			id: "view",
			title: "View",
			groups: [
				{
					title: "Zoom",
					items: [
						{ kind: "button", icon: "zoom-out", tooltip: "Zoom out", onClick: host.zoomOut },
						{ kind: "button", icon: "zoom-in", tooltip: "Zoom in", onClick: host.zoomIn },
						{ kind: "button", icon: "maximize", tooltip: "Fit to pane", onClick: host.zoomToFit },
					],
				},
				{
					title: "Show",
					items: [
						{
							kind: "button",
							icon: "sticky-note",
							label: "Notes",
							tooltip: "Toggle speaker notes",
							isActive: host.notesShown,
							onClick: host.toggleNotes,
						},
						{
							kind: "button",
							icon: "panel-left",
							label: "Thumbnails",
							tooltip: "Toggle the thumbnail rail",
							onClick: host.toggleThumbnails,
						},
					],
				},
				{
					title: "Export",
					items: [
						{
							kind: "button",
							icon: "image-down",
							label: "PNG",
							tooltip: "Export this slide as a PNG",
							onClick: host.exportPng,
						},
						{
							kind: "button",
							icon: "file-text",
							label: "Markdown",
							tooltip: "Extract the deck's text to a note",
							onClick: host.extractMarkdown,
						},
						{
							kind: "button",
							icon: "external-link",
							tooltip: "Open in the default app",
							onClick: host.openExternally,
						},
					],
				},
			],
		},
	];
}
