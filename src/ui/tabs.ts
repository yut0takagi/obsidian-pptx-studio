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
	changeShape,
	copyFormatting,
	fillState,
	flipSelection,
	geometryState,
	hasCopiedFormat,
	outlineState,
	pasteFormatting,
	rotateBy,
	setFill,
	setGeometry,
	setOutline,
	setVerticalAnchor,
	textState,
} from "../edit/formatCommands";
import {
	type TableSelection,
	deleteTableColumns,
	deleteTableRows,
	hasTableSelection,
	insertTableColumn,
	insertTableRow,
	mergeTableCells,
	setCellFill,
	splitTableCells,
} from "../edit/tableCommands";
import { insertAutoShape, insertLine, insertTextBox } from "../edit/insertCommands";
import {
	canDeleteSlide,
	deleteCurrentSlide,
	duplicateCurrentSlide,
	moveCurrentSlide,
	newSlide,
	setSlideBackgroundColor,
} from "../edit/slideCommands";
import { type StringKey, t } from "../i18n";
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
	pickHyperlink: () => void;
	tableSelection: TableSelection;
	slideBackground: () => string | null;
	exportPng: () => void;
	extractMarkdown: () => void;
	openExternally: () => void;
}

/** Presets offered by the shapes menu, in the order PowerPoint groups them. */
export const SHAPE_PRESETS: { preset: string; key: StringKey }[] = [
	{ preset: "rect", key: "shape.rect" },
	{ preset: "roundRect", key: "shape.roundRect" },
	{ preset: "ellipse", key: "shape.ellipse" },
	{ preset: "triangle", key: "shape.triangle" },
	{ preset: "diamond", key: "shape.diamond" },
	{ preset: "pentagon", key: "shape.pentagon" },
	{ preset: "hexagon", key: "shape.hexagon" },
	{ preset: "star5", key: "shape.star5" },
	{ preset: "rightArrow", key: "shape.rightArrow" },
	{ preset: "chevron", key: "shape.chevron" },
	{ preset: "flowChartDecision", key: "shape.flowChartDecision" },
	{ preset: "wedgeRectCallout", key: "shape.wedgeRectCallout" },
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

/** A compact labelled number input, used for position and size. */
function numberField(
	label: string,
	tooltip: string,
	value: () => number | null,
	onChange: (value: number) => void,
): RibbonItem {
	return { kind: "number", label, tooltip, width: "4.5em", value, onChange };
}

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
			tooltip: t("cmd.font"),
			width: "9.5em",
			options: () => FONTS.map((f) => ({ value: f, label: f })),
			value: () => state()?.font ?? "",
			isEnabled: () => host.canEdit() && hasSelection(),
			onChange: (value) => host.run((ctx) => applyTextFormat(ctx, { font: value }, t("cmd.font"))),
		},
		{
			kind: "select",
			tooltip: t("cmd.fontSize"),
			width: "4.5em",
			options: () => FONT_SIZES.map((s) => ({ value: String(s), label: String(s) })),
			value: () => {
				const size = state()?.size;
				return size === null || size === undefined ? "" : String(Math.round(size));
			},
			isEnabled: () => host.canEdit() && hasSelection(),
			onChange: (value) =>
				host.run((ctx) => applyTextFormat(ctx, { size: Number(value) }, t("cmd.fontSize"))),
		},
		textButton("bold", t("cmd.bold"), "bold"),
		textButton("italic", t("cmd.italic"), "italic"),
		textButton("underline", t("cmd.underline"), "underline"),
		textButton("strikethrough", t("cmd.strike"), "strike"),
		{
			kind: "color",
			icon: "baseline",
			tooltip: t("cmd.textColour"),
			value: () => state()?.color ?? null,
			isEnabled: () => host.canEdit() && hasSelection(),
			onChange: (color) => host.run((ctx) => applyTextFormat(ctx, { color }, t("cmd.textColour"))),
		},
		{
			kind: "button",
			icon: "link",
			tooltip: t("cmd.hyperlink"),
			isEnabled: () => host.canEdit() && hasSelection(),
			onClick: () => host.pickHyperlink(),
		},
	];

	const paragraphGroup: RibbonItem[] = [
		{
			kind: "button",
			icon: "list",
			tooltip: t("cmd.bulletList"),
			isEnabled: () => host.canEdit() && hasSelection(),
			isActive: () => state()?.bulleted ?? false,
			onClick: () =>
				host.run((ctx) =>
					applyParagraphFormat(
						ctx,
						{ bullet: state()?.bulleted ? "none" : "char" },
						t("cmd.bulletList"),
					),
				),
		},
		{
			kind: "button",
			icon: "list-ordered",
			tooltip: t("cmd.numberedList"),
			isEnabled: () => host.canEdit() && hasSelection(),
			onClick: () => host.run((ctx) => applyParagraphFormat(ctx, { bullet: "number" }, t("cmd.numberedList"))),
		},
		{
			kind: "button",
			icon: "indent-decrease",
			tooltip: t("cmd.outdent"),
			isEnabled: () => host.canEdit() && hasSelection(),
			onClick: () => host.run((ctx) => applyParagraphFormat(ctx, { levelDelta: -1 }, t("cmd.outdent"))),
		},
		{
			kind: "button",
			icon: "indent-increase",
			tooltip: t("cmd.indent"),
			isEnabled: () => host.canEdit() && hasSelection(),
			onClick: () => host.run((ctx) => applyParagraphFormat(ctx, { levelDelta: 1 }, t("cmd.indent"))),
		},
		{ kind: "separator" },
		alignButton("align-left", t("cmd.alignLeft"), "l"),
		alignButton("align-center", t("cmd.alignCentre"), "ctr"),
		alignButton("align-right", t("cmd.alignRight"), "r"),
		alignButton("align-justify", t("cmd.justify"), "just"),
		{
			kind: "menu",
			icon: "chevrons-up-down",
			tooltip: t("cmd.spacingMenu"),
			isEnabled: () => host.canEdit() && hasSelection(),
			build: (menu: Menu) => {
				for (const value of [1, 1.15, 1.5, 2]) {
					menu.addItem((item) =>
						item
							.setTitle(t("cmd.lineSpacing", { value }))
							.onClick(() =>
								host.run((ctx) => applyParagraphFormat(ctx, { lineSpacing: value }, t("cmd.lineSpacing", { value }))),
							),
					);
				}
				menu.addSeparator();
				const anchors: { label: string; value: "t" | "ctr" | "b" }[] = [
					{ label: t("cmd.anchorTop"), value: "t" },
					{ label: t("cmd.anchorMiddle"), value: "ctr" },
					{ label: t("cmd.anchorBottom"), value: "b" },
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
			label: t("cmd.arrange"),
			tooltip: t("cmd.arrangeTooltip"),
			isEnabled: () => host.canEdit() && hasSelection(),
			build: (menu: Menu) => {
				const options = [
					{ label: t("cmd.bringToFront"), mode: "front" as const, icon: "bring-to-front" },
					{ label: t("cmd.bringForward"), mode: "forward" as const, icon: "arrow-up" },
					{ label: t("cmd.sendBackward"), mode: "backward" as const, icon: "arrow-down" },
					{ label: t("cmd.sendToBack"), mode: "back" as const, icon: "send-to-back" },
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
			label: t("cmd.align"),
			tooltip: t("cmd.alignTooltip"),
			isEnabled: () => host.canEdit() && hasSelection(),
			build: (menu: Menu) => {
				const options = [
					{ label: t("cmd.alignLeft"), mode: "left" as const },
					{ label: t("cmd.alignCentre"), mode: "centerH" as const },
					{ label: t("cmd.alignRight"), mode: "right" as const },
					{ label: t("cmd.alignTop"), mode: "top" as const },
					{ label: t("cmd.alignMiddle"), mode: "middle" as const },
					{ label: t("cmd.alignBottom"), mode: "bottom" as const },
				];
				for (const option of options) {
					menu.addItem((item) =>
						item.setTitle(option.label).onClick(() => host.run((ctx) => alignSelection(ctx, option.mode))),
					);
				}
				menu.addSeparator();
				menu.addItem((item) =>
					item
						.setTitle(t("cmd.distributeH"))
						.setDisabled(selectionCount() < 3)
						.onClick(() => host.run((ctx) => distributeSelection(ctx, "h"))),
				);
				menu.addItem((item) =>
					item
						.setTitle(t("cmd.distributeV"))
						.setDisabled(selectionCount() < 3)
						.onClick(() => host.run((ctx) => distributeSelection(ctx, "v"))),
				);
			},
		},
		{
			kind: "button",
			icon: "group",
			tooltip: t("cmd.group"),
			isEnabled: () => host.canEdit() && selectionCount() > 1,
			onClick: () => host.run(groupSelection),
		},
		{
			kind: "button",
			icon: "ungroup",
			tooltip: t("cmd.ungroup"),
			isEnabled: () => host.canEdit() && hasGroup(),
			onClick: () => host.run(ungroupSelection),
		},
	];

	const shapeStyleItems: RibbonItem[] = [
		{
			kind: "color",
			icon: "paint-bucket",
			tooltip: t("cmd.shapeFill"),
			allowNone: true,
			value: () => fillState(host.ctx()),
			isEnabled: () => host.canEdit() && hasSelection(),
			onChange: (color) => host.run((ctx) => setFill(ctx, color)),
		},
		{
			kind: "color",
			icon: "square",
			tooltip: t("cmd.shapeOutline"),
			allowNone: true,
			value: () => outlineState(host.ctx()),
			isEnabled: () => host.canEdit() && hasSelection(),
			onChange: (color) => host.run((ctx) => setOutline(ctx, { color }, t("cmd.shapeOutline"))),
		},
		{
			kind: "menu",
			icon: "pencil-ruler",
			tooltip: t("cmd.outlineStyle"),
			isEnabled: () => host.canEdit() && hasSelection(),
			build: (menu: Menu) => {
				for (const width of [0.75, 1, 1.5, 2.25, 3, 4.5, 6]) {
					menu.addItem((item) =>
						item
							.setTitle(t("cmd.outlineWeight", { value: width }))
							.onClick(() => host.run((ctx) => setOutline(ctx, { width }, t("cmd.outlineStyle")))),
					);
				}
				menu.addSeparator();
				const dashes = [
					{ label: t("cmd.dashSolid"), value: "solid" },
					{ label: t("cmd.dashDashed"), value: "dash" },
					{ label: t("cmd.dashDotted"), value: "sysDot" },
				];
				for (const dash of dashes) {
					menu.addItem((item) =>
						item
							.setTitle(dash.label)
							.onClick(() => host.run((ctx) => setOutline(ctx, { dash: dash.value }, t("cmd.outlineStyle")))),
					);
				}
			},
		},
	];

	const shapesMenu: RibbonItem = {
		kind: "menu",
		icon: "shapes",
		label: t("cmd.shapes"),
		tooltip: t("cmd.insertShape"),
		isEnabled: host.canEdit,
		build: (menu: Menu) => {
			for (const shape of SHAPE_PRESETS) {
				menu.addItem((item) =>
					item
						.setTitle(t(shape.key))
						.onClick(() => host.run((ctx) => insertAutoShape(ctx, shape.preset, t(shape.key)))),
				);
			}
			menu.addSeparator();
			menu.addItem((item) => item.setTitle(t("cmd.line")).onClick(() => host.run(insertLine)));
		},
	};

	const clipboardItems: RibbonItem[] = [
		{
			kind: "button",
			icon: "clipboard-paste",
			tooltip: t("cmd.paste"),
			isEnabled: () => host.canEdit() && hasClipboard(),
			onClick: () => host.run((ctx) => pasteClipboard(ctx)),
		},
		{
			kind: "button",
			icon: "scissors",
			tooltip: t("cmd.cut"),
			isEnabled: () => host.canEdit() && hasSelection(),
			onClick: () => host.run(cutSelection),
		},
		{
			kind: "button",
			icon: "copy",
			tooltip: t("cmd.copy"),
			isEnabled: () => host.canEdit() && hasSelection(),
			onClick: () => host.run(copySelection),
		},
		{
			kind: "button",
			icon: "copy-plus",
			tooltip: t("cmd.duplicate"),
			isEnabled: () => host.canEdit() && hasSelection(),
			onClick: () => host.run(duplicateSelection),
		},
	];

	const slideItems: RibbonItem[] = [
		{
			kind: "button",
			icon: "file-plus",
			label: t("cmd.newSlideShort"),
			tooltip: t("cmd.newSlide"),
			isEnabled: host.canEdit,
			onClick: () => host.runSlide((ctx) => newSlide(ctx)),
		},
		{
			kind: "button",
			icon: "layout-template",
			tooltip: t("cmd.newSlideFromLayout"),
			isEnabled: host.canEdit,
			onClick: () => host.pickLayout(),
		},
		{
			kind: "button",
			icon: "files",
			tooltip: t("cmd.duplicateSlide"),
			isEnabled: host.canEdit,
			onClick: () => host.runSlide(duplicateCurrentSlide),
		},
		{
			kind: "button",
			icon: "trash-2",
			tooltip: t("cmd.deleteSlide"),
			isEnabled: () => {
				const ctx = host.ctx();
				return host.canEdit() && ctx !== null && canDeleteSlide(ctx);
			},
			onClick: () => host.run(deleteCurrentSlide),
		},
		{
			kind: "button",
			icon: "arrow-up",
			tooltip: t("cmd.moveSlideUp"),
			isEnabled: host.canEdit,
			onClick: () => host.runSlide((ctx) => moveCurrentSlide(ctx, -1)),
		},
		{
			kind: "button",
			icon: "arrow-down",
			tooltip: t("cmd.moveSlideDown"),
			isEnabled: host.canEdit,
			onClick: () => host.runSlide((ctx) => moveCurrentSlide(ctx, 1)),
		},
	];

	return [
		{
			id: "home",
			title: t("tab.home"),
			groups: [
				{
					title: t("group.undo"),
					items: [
						{
							kind: "button",
							icon: "undo-2",
							tooltip: t("cmd.undo"),
							isEnabled: host.canUndo,
							onClick: host.undo,
						},
						{
							kind: "button",
							icon: "redo-2",
							tooltip: t("cmd.redo"),
							isEnabled: host.canRedo,
							onClick: host.redo,
						},
						{
							kind: "button",
							icon: "save",
							tooltip: t("cmd.save"),
							isEnabled: host.isDirty,
							onClick: host.save,
						},
					],
				},
				{ title: t("group.clipboard"), items: clipboardItems },
				{ title: t("group.slides"), items: slideItems.slice(0, 4) },
				{ title: t("group.font"), items: fontGroup },
				{ title: t("group.paragraph"), items: paragraphGroup },
				{ title: t("group.drawing"), items: [shapesMenu, ...shapeStyleItems, ...arrangeItems] },
				{
					title: t("group.editing"),
					items: [
						{
							kind: "button",
							icon: "mouse-pointer-2",
							tooltip: t("cmd.selectAll"),
							isEnabled: host.canEdit,
							onClick: host.selectAll,
						},
						{
							kind: "button",
							icon: "trash",
							tooltip: t("cmd.delete"),
							isEnabled: () => host.canEdit() && hasSelection(),
							onClick: () => host.run(deleteSelection),
						},
					],
				},
			],
		},
		{
			id: "insert",
			title: t("tab.insert"),
			groups: [
				{ title: t("group.slides"), items: slideItems },
				{
					title: t("group.objects"),
					items: [
						{
							kind: "button",
							icon: "type",
							label: t("cmd.textBox"),
							tooltip: t("cmd.insertTextBox"),
							isEnabled: host.canEdit,
							onClick: () => host.run(insertTextBox),
						},
						shapesMenu,
						{
							kind: "button",
							icon: "image",
							label: t("cmd.picture"),
							tooltip: t("cmd.insertPicture"),
							isEnabled: host.canEdit,
							onClick: () => host.pickImage(),
						},
						{
							kind: "button",
							icon: "table",
							label: t("cmd.table"),
							tooltip: t("cmd.insertTable"),
							isEnabled: host.canEdit,
							onClick: () => host.pickTable(),
						},
					],
				},
			],
		},
		{
			id: "format",
			title: t("tab.format"),
			groups: [
				{ title: t("group.shapeStyles"), items: shapeStyleItems },
				{
					title: t("group.shape"),
					items: [
						{
							kind: "menu",
							icon: "shapes",
							label: t("cmd.changeShape"),
							tooltip: t("cmd.changeShapeTooltip"),
							isEnabled: () => host.canEdit() && hasSelection(),
							build: (menu: Menu) => {
								for (const shape of SHAPE_PRESETS) {
									menu.addItem((item) =>
										item
											.setTitle(t(shape.key))
											.onClick(() => host.run((ctx) => changeShape(ctx, shape.preset, t(shape.key)))),
									);
								}
							},
						},
						{
							kind: "button",
							icon: "paintbrush",
							tooltip: t("cmd.copyFormat"),
							isEnabled: () => host.canEdit() && hasSelection(),
							onClick: () => host.run(copyFormatting),
						},
						{
							kind: "button",
							icon: "clipboard-check",
							tooltip: t("cmd.pasteFormat"),
							isEnabled: () => host.canEdit() && hasSelection() && hasCopiedFormat(),
							onClick: () => host.run(pasteFormatting),
						},
					],
				},
				{ title: t("group.arrange"), items: arrangeItems },
				{
					title: t("group.rotate"),
					items: [
						{
							kind: "button",
							icon: "rotate-cw",
							tooltip: t("cmd.rotateRight"),
							isEnabled: () => host.canEdit() && hasSelection(),
							onClick: () => host.run((ctx) => rotateBy(ctx, 90)),
						},
						{
							kind: "button",
							icon: "rotate-ccw",
							tooltip: t("cmd.rotateLeft"),
							isEnabled: () => host.canEdit() && hasSelection(),
							onClick: () => host.run((ctx) => rotateBy(ctx, -90)),
						},
						{
							kind: "button",
							icon: "flip-horizontal",
							tooltip: t("cmd.flipH"),
							isEnabled: () => host.canEdit() && hasSelection(),
							onClick: () => host.run((ctx) => flipSelection(ctx, "h")),
						},
						{
							kind: "button",
							icon: "flip-vertical",
							tooltip: t("cmd.flipV"),
							isEnabled: () => host.canEdit() && hasSelection(),
							onClick: () => host.run((ctx) => flipSelection(ctx, "v")),
						},
					],
				},
				{
					title: t("group.positionSize"),
					items: [
						numberField("X", t("cmd.posX"), () => geometryState(host.ctx())?.x ?? null, (x) => host.run((ctx) => setGeometry(ctx, { x }, "Set position"))),
						numberField("Y", t("cmd.posY"), () => geometryState(host.ctx())?.y ?? null, (y) => host.run((ctx) => setGeometry(ctx, { y }, "Set position"))),
						numberField("W", t("cmd.sizeW"), () => geometryState(host.ctx())?.w ?? null, (w) => host.run((ctx) => setGeometry(ctx, { w }, "Set size"))),
						numberField("H", t("cmd.sizeH"), () => geometryState(host.ctx())?.h ?? null, (h) => host.run((ctx) => setGeometry(ctx, { h }, "Set size"))),
						numberField("°", t("cmd.rotation"), () => geometryState(host.ctx())?.rotation ?? null, (rotation) => host.run((ctx) => setGeometry(ctx, { rotation }, "Set rotation"))),
					],
				},
			],
		},
		{
			id: "design",
			title: t("tab.design"),
			groups: [
				{
					title: t("group.background"),
					items: [
						{
							kind: "color",
							icon: "paint-bucket",
							tooltip: t("cmd.slideBackground"),
							allowNone: true,
							value: host.slideBackground,
							isEnabled: host.canEdit,
							onChange: (color) => host.run((ctx) => setSlideBackgroundColor(ctx, color)),
						},
					],
				},
				{ title: t("group.slides"), items: slideItems },
			],
		},
		{
			id: "table",
			title: t("tab.table"),
			visible: () => hasTableSelection(host.ctx(), host.tableSelection),
			groups: [
				{
					title: t("group.rowsColumns"),
					items: [
						{
							kind: "button",
							icon: "between-vertical-start",
							label: t("cmd.rowAbove"),
							tooltip: t("cmd.rowAboveTooltip"),
							onClick: () => host.run((ctx) => insertTableRow(ctx, host.tableSelection, "above")),
						},
						{
							kind: "button",
							icon: "between-vertical-end",
							label: t("cmd.rowBelow"),
							tooltip: t("cmd.rowBelowTooltip"),
							onClick: () => host.run((ctx) => insertTableRow(ctx, host.tableSelection, "below")),
						},
						{
							kind: "button",
							icon: "between-horizontal-start",
							label: t("cmd.columnLeft"),
							tooltip: t("cmd.columnLeftTooltip"),
							onClick: () => host.run((ctx) => insertTableColumn(ctx, host.tableSelection, "left")),
						},
						{
							kind: "button",
							icon: "between-horizontal-end",
							label: t("cmd.columnRight"),
							tooltip: t("cmd.columnRightTooltip"),
							onClick: () => host.run((ctx) => insertTableColumn(ctx, host.tableSelection, "right")),
						},
						{ kind: "separator" },
						{
							kind: "button",
							icon: "rows-3",
							tooltip: t("cmd.deleteRows"),
							onClick: () => host.run((ctx) => deleteTableRows(ctx, host.tableSelection)),
						},
						{
							kind: "button",
							icon: "columns-3",
							tooltip: t("cmd.deleteColumns"),
							onClick: () => host.run((ctx) => deleteTableColumns(ctx, host.tableSelection)),
						},
					],
				},
				{
					title: t("group.cells"),
					items: [
						{
							kind: "button",
							icon: "table-cells-merge",
							label: t("cmd.merge"),
							tooltip: t("cmd.mergeTooltip"),
							onClick: () => host.run((ctx) => mergeTableCells(ctx, host.tableSelection)),
						},
						{
							kind: "button",
							icon: "table-cells-split",
							label: t("cmd.split"),
							tooltip: t("cmd.splitTooltip"),
							onClick: () => host.run((ctx) => splitTableCells(ctx, host.tableSelection)),
						},
						{
							kind: "color",
							icon: "paint-bucket",
							tooltip: t("cmd.cellFill"),
							allowNone: true,
							value: () => null,
							onChange: (color) => host.run((ctx) => setCellFill(ctx, host.tableSelection, color)),
						},
					],
				},
			],
		},
		{
			id: "view",
			title: t("tab.view"),
			groups: [
				{
					title: t("group.zoom"),
					items: [
						{ kind: "button", icon: "zoom-out", tooltip: t("zoom.out"), onClick: host.zoomOut },
						{ kind: "button", icon: "zoom-in", tooltip: t("zoom.in"), onClick: host.zoomIn },
						{ kind: "button", icon: "maximize", tooltip: t("zoom.fit"), onClick: host.zoomToFit },
					],
				},
				{
					title: t("group.show"),
					items: [
						{
							kind: "button",
							icon: "sticky-note",
							label: t("cmd.notes"),
							tooltip: t("cmd.notesTooltip"),
							isActive: host.notesShown,
							onClick: host.toggleNotes,
						},
						{
							kind: "button",
							icon: "panel-left",
							label: t("cmd.thumbnails"),
							tooltip: t("cmd.thumbnailsTooltip"),
							onClick: host.toggleThumbnails,
						},
					],
				},
				{
					title: t("group.export"),
					items: [
						{
							kind: "button",
							icon: "image-down",
							label: t("cmd.exportPng"),
							tooltip: t("cmd.exportPngTooltip"),
							onClick: host.exportPng,
						},
						{
							kind: "button",
							icon: "file-text",
							label: t("cmd.exportMarkdown"),
							tooltip: t("cmd.exportMarkdownTooltip"),
							onClick: host.extractMarkdown,
						},
						{
							kind: "button",
							icon: "external-link",
							tooltip: t("view.openExternally"),
							onClick: host.openExternally,
						},
					],
				},
			],
		},
	];
}
