import { FileView, Menu, Notice, type TFile, type WorkspaceLeaf } from "obsidian";
import type { LoadedDeck } from "../DeckCache";
import { relsPathFor } from "../ooxml/rels";
import { DEFAULT_SETTINGS } from "../settings";
import type { Deck } from "../pptx/types";
import { DeckEditor } from "../edit/DeckEditor";
import { EditController } from "../edit/EditController";
import { Selection } from "../edit/Selection";
import { CropController } from "../edit/CropController";
import { GuideController } from "../edit/GuideController";
import { ShapeEditor } from "../edit/ShapeEditor";
import {
	type CommandContext,
	copySelection,
	cutSelection,
	deleteSelection,
	duplicateSelection,
	groupSelection,
	hasClipboard,
	pasteClipboard,
	reorderSelection,
	selectedShapes,
	ungroupSelection,
} from "../edit/commands";
import { applyParagraphFormatAt, hyperlinkState, setHyperlink } from "../edit/formatCommands";
import { insertPicture, insertTable } from "../edit/insertCommands";
import { TableSelection } from "../edit/tableCommands";
import { ConflictError, saveDeck } from "../edit/save";
import {
	canDeleteSlide,
	deleteCurrentSlide,
	duplicateCurrentSlide,
	listLayouts,
	moveCurrentSlide,
	newSlide,
	reorderSlide,
} from "../edit/slideCommands";
import { rebuildDeck, rebuildSlideAt } from "../pptx/parse";
import { DeckViewer } from "../render/DeckViewer";
import { ContextToolbar } from "../ui/ContextToolbar";
import { Ribbon } from "../ui/Ribbon";
import { InspectorPane } from "../ui/InspectorPane";
import { PaneSplitter } from "../ui/PaneSplitter";
import { SelectionPane } from "../ui/SelectionPane";
import {
	FindReplaceModal,
	ImagePickerModal,
	LayoutPickerModal,
	PromptModal,
	TableSizeModal,
	imageDimensions,
} from "../ui/modals";
import { type RibbonHost, buildTabs } from "../ui/tabs";
import { t } from "../i18n";
import type PptxStudioPlugin from "../main";
import { chooseImage, clipboardImageName } from "../edit/clipboardImage";
import { findMatches, replaceAll } from "../edit/findReplace";

export const VIEW_TYPE_PPTX = "pptx-studio";

/** Parts that say nothing about the model's structure, whatever else changed. */
const NEUTRAL = ["ppt/media/", "[Content_Types].xml", "ppt/viewProps.xml"];

/**
 * The index of the one slide an edit touched, or null when it reached wider.
 *
 * A slide's own part and its relationships belong to it alone; the
 * presentation, the layouts and the masters belong to everything.
 */
function singleSlideScope(deck: Deck, parts: string[]): number | null {
	let found: number | null = null;
	for (const part of parts) {
		if (NEUTRAL.some((prefix) => part.startsWith(prefix))) continue;
		const index = deck.slides.findIndex(
			(slide) => part === slide.partPath || part === relsPathFor(slide.partPath),
		);
		if (index === -1) return null;
		if (found !== null && found !== index) return null;
		found = index;
	}
	return found;
}

/** The full-tab deck editor: ribbon, slide canvas, selection, save. */
export class PptxView extends FileView {
	private viewer: DeckViewer | null = null;
	private ribbon: Ribbon | null = null;
	private editController: EditController | null = null;
	private shapeEditor: ShapeEditor | null = null;
	private guides: GuideController | null = null;
	private crop: CropController | null = null;
	private showRulers = true;
	private showSelectionPane = false;
	private selectionPane: SelectionPane | null = null;
	private inspectorPane: InspectorPane | null = null;
	private sideSplitter: PaneSplitter | null = null;
	private contextToolbar: ContextToolbar | null = null;
	private uiFrame = 0;
	private deckEditor: DeckEditor | null = null;
	private readonly selection = new Selection();
	private readonly tableSelection = new TableSelection();
	private activeSlideEl: HTMLElement | null = null;
	private loaded: LoadedDeck | null = null;
	/** mtime the open deck was read from, used to detect outside edits. */
	private baseMtime = 0;
	private showThumbnails = true;

	constructor(
		leaf: WorkspaceLeaf,
		private readonly plugin: PptxStudioPlugin,
	) {
		super(leaf);
		this.selection.onChange(() => {
			this.tableSelection.retain(this.selection.ids);
			this.scheduleUi();
		});
		this.tableSelection.onChange(() => {
			this.tableSelection.paint(this.activeSlideEl);
			this.scheduleUi();
		});
	}

	getViewType(): string {
		return VIEW_TYPE_PPTX;
	}

	getIcon(): string {
		return "presentation";
	}

	getDisplayText(): string {
		return this.file?.basename ?? t("view.title");
	}

	canAcceptExtension(extension: string): boolean {
		return extension === "pptx";
	}

	async onLoadFile(file: TFile): Promise<void> {
		this.teardown();
		this.contentEl.empty();
		this.contentEl.addClass("pptx-view");
		this.showThumbnails = this.plugin.settings.showThumbnails;

		const loading = this.contentEl.createDiv({ cls: "pptx-message", text: t("view.loading") });
		try {
			const loaded = await this.plugin.decks.get(file);
			loading.remove();
			this.loaded = loaded;
			this.baseMtime = file.stat.mtime;
			this.build();
		} catch (error) {
			loading.remove();
			this.renderError(file, error as Error);
		}
	}

	async onUnloadFile(_file: TFile): Promise<void> {
		this.editController?.commit();
		this.teardown();
		this.contentEl.empty();
	}

	onunload(): void {
		this.editController?.commit();
		this.teardown();
	}

	// --------------------------------------------------------------- build

	private build(): void {
		const loaded = this.loaded;
		if (!loaded) return;

		this.deckEditor = new DeckEditor(loaded.pkg, {
			onChanged: (rebuild, parts) => this.afterEdit(rebuild, parts),
		});

		this.editController = new EditController({
			isEnabled: () => true,
			editor: this.deckEditor,
			onCancelled: () => this.viewer?.invalidate(),
			onReturnToShape: () => this.viewer?.focus(),
			onModeChange: (editing) => {
				this.shapeEditor?.setEditing(editing);
				this.scheduleUi();
			},
			onListLevel: (delta, target) =>
				this.run((ctx) =>
					applyParagraphFormatAt(
						ctx,
						target.shapeId,
						target.paragraph,
						{ levelDelta: delta },
						delta > 0 ? "Indent" : "Outdent",
					),
				),
		});

		this.guides = new GuideController({
			pkg: loaded.pkg,
			editor: this.deckEditor,
			getScale: () => this.viewer?.currentScale() ?? 1,
			getSlideSize: () => ({ width: loaded.deck.width, height: loaded.deck.height }),
			isEnabled: () => !this.editController?.isEditing,
		});

		this.crop = new CropController({
			getContext: () => this.context(),
			getScale: () => this.viewer?.currentScale() ?? 1,
			onChanged: () => this.scheduleUi(),
		});

		this.shapeEditor = new ShapeEditor({
			selection: this.selection,
			editor: this.deckEditor,
			// Dragging is off while a text box is open, so a stray drag inside the
			// caret's box cannot move the shape out from under the editor.
			isEnabled: () => !this.editController?.isEditing,
			getScale: () => this.viewer?.currentScale() ?? 1,
			getContext: () => this.context(),
			onContextMenu: (event) => this.showContextMenu(event),
			onEditText: (shapeId) => this.editTextOf(shapeId),
			onTypeText: (shapeId, text) => this.replaceTextOf(shapeId, text),
			// The border, the handles and the other shapes all belong to the box
			// rather than to its text, so a press on one ends the text edit first.
			onLeaveText: () => {
				this.editController?.commit();
				this.viewer?.focus();
			},
			onCellPointerDown: (shape, row, column, additive) =>
				this.tableSelection.select(shape.id, row, column, additive),
			extraGuides: () => this.guides?.snapTargets() ?? { xs: [], ys: [] },
			// Cropping owns the pointer while it is on, then guides, then shapes.
			claimPointer: (event) =>
				(this.crop?.tryGrab(event) ?? false) || (this.guides?.tryGrab(event) ?? false),
		});

		this.contextToolbar = new ContextToolbar({
			getContext: () => this.context(),
			selection: this.selection,
			tableSelection: this.tableSelection,
			run: (fn) => this.run(fn),
			getSlideEl: () => this.activeSlideEl,
			getViewportEl: () => this.viewer?.stageElement ?? null,
			getScale: () => this.viewer?.currentScale() ?? 1,
			isEditing: () => this.editController?.isEditing ?? false,
			canCrop: () => this.crop?.canCrop() ?? false,
			cropActive: () => this.crop?.active ?? false,
			toggleCrop: () => this.crop?.toggle(),
		});

		this.ribbon = new Ribbon(this.contentEl, buildTabs(this.ribbonHost()), {
			collapsed: this.plugin.settings.ribbonCollapsed,
			initialTab: this.plugin.settings.ribbonTab,
			onStateChange: ({ collapsed, tab }) => {
				if (
					this.plugin.settings.ribbonCollapsed === collapsed &&
					this.plugin.settings.ribbonTab === tab
				) {
					return;
				}
				this.plugin.settings.ribbonCollapsed = collapsed;
				this.plugin.settings.ribbonTab = tab;
				void this.plugin.saveSettings();
			},
		});
		this.createViewer(1);
		this.contentEl.addEventListener("keydown", this.onKeyDown, { capture: true });
		this.contentEl.addEventListener("paste", this.onPaste, { capture: true });
		this.viewer?.focus();
		this.scheduleUi();
	}

	private createViewer(slideNumber: number): void {
		const loaded = this.loaded;
		if (!loaded) return;
		const settings = this.plugin.settings;
		this.viewer = new DeckViewer(this.contentEl, {
			deck: loaded.deck,
			pkg: loaded.pkg,
			compact: false,
			chrome: "status",
			showThumbnails: this.showThumbnails,
			showNotes: settings.showNotes,
			fitMode: settings.fitMode,
			editor: this.editController ?? undefined,
			shapeEditor: this.shapeEditor ?? undefined,
			onSlideChange: () => {
				this.selection.clear();
				this.tableSelection.clear();
			},
			showRulers: this.showRulers,
			sidePane: this.showSelectionPane,
			sidePaneWidth: settings.sidePaneWidth,
			onSidePaneWidth: (width) => {
				if (settings.sidePaneWidth === width) return;
				settings.sidePaneWidth = width;
				void this.plugin.saveSettings();
			},
			onRulerDrag: (orientation, position, event) =>
				this.guides?.beginFromRuler(event, orientation, position),
			onViewportChanged: () => this.contextToolbar?.refresh(),
			onRendered: (slideEl) => {
				this.activeSlideEl = slideEl;
				this.tableSelection.paint(slideEl);
				this.guides?.setActive(slideEl);
				this.crop?.setActive(slideEl);
				this.contextToolbar?.refresh();
			},
			onThumbnailMenu: (index, event) => this.showSlideMenu(index, event),
			onReorder: (from, to) => {
				this.run((ctx) => reorderSlide(ctx, from, to));
				this.viewer?.go(to);
			},
		});
		if (slideNumber > 1) this.viewer.go(slideNumber - 1);
		this.viewer.setDirty(loaded.pkg.isDirty);

		const host = this.viewer.sidePane;
		this.selectionPane = host
			? new SelectionPane(host, {
					getContext: () => this.context(),
					selection: this.selection,
					run: (fn) => this.run(fn),
					collapsed: settings.shapeListCollapsed,
					onCollapsed: (collapsed) => {
						settings.shapeListCollapsed = collapsed;
						void this.plugin.saveSettings();
						this.updateSideLayout();
					},
				})
			: null;
		// The divider goes between the two sections, so it is built between them.
		this.sideSplitter =
			host && this.selectionPane
				? new PaneSplitter(host, {
						above: this.selectionPane.element,
						height: settings.shapeListHeight,
						defaultHeight: DEFAULT_SETTINGS.shapeListHeight,
						onHeight: (height) => {
							if (settings.shapeListHeight === height) return;
							settings.shapeListHeight = height;
							void this.plugin.saveSettings();
						},
					})
				: null;
		this.inspectorPane = host
			? new InspectorPane(host, {
					getContext: () => this.context(),
					run: (fn) => this.run(fn),
					collapsed: settings.inspectorCollapsed,
					onCollapsed: (collapsed) => {
						settings.inspectorCollapsed = collapsed;
						void this.plugin.saveSettings();
						this.updateSideLayout();
					},
				})
			: null;
		this.updateSideLayout();
	}

	/**
	 * Who gets the space in the side pane.
	 *
	 * With both sections open the divider decides, and the shape list holds the
	 * height it was dragged to. Fold one away and the question stops being a
	 * question: the other takes everything that is left, and the divider goes.
	 */
	private updateSideLayout(): void {
		const settings = this.plugin.settings;
		const list = this.selectionPane?.element;
		if (!list) return;
		const bothOpen = !settings.shapeListCollapsed && !settings.inspectorCollapsed;
		list.toggleClass("is-flexible", !bothOpen && !settings.shapeListCollapsed);
		this.sideSplitter?.setEnabled(bothOpen);
		if (bothOpen) this.sideSplitter?.setHeight(settings.shapeListHeight);
		else list.setCssStyles({ height: "" });
	}

	/** Rebuild the model and repaint after an edit. */
	private afterEdit(rebuild: boolean, parts: string[] = []): void {
		const loaded = this.loaded;
		if (!loaded) return;
		if (rebuild) {
			// Re-derive only the slide that changed when that is all that changed.
			const only = singleSlideScope(loaded.deck, parts);
			if (only !== null && rebuildSlideAt(loaded.pkg, loaded.deck, only)) {
				this.viewer?.refreshSlide(only);
			} else {
				loaded.deck = rebuildDeck(loaded.pkg, this.file?.basename ?? "Deck");
				this.viewer?.setDeck(loaded.deck);
			}
			// Shapes can vanish from under a selection: deleted, ungrouped, undone.
			const slide = loaded.deck.slides[this.selection.slideIndex];
			const existing = new Set((slide?.shapes ?? []).filter((s) => s.source).map((s) => s.id));
			this.selection.retain(existing);
			this.tableSelection.retain(this.selection.ids);
			this.shapeEditor?.refresh();
			this.tableSelection.paint(this.activeSlideEl);
			this.guides?.reload();
		}
		this.viewer?.setDirty(true);
		this.scheduleUi();
	}

	/**
	 * Coalesce the chrome refresh into one frame.
	 *
	 * A single edit can otherwise ask the ribbon, the selection pane, the
	 * floating toolbar and the status bar to re-evaluate several times over,
	 * which is exactly the sort of work that turns a keystroke into a stutter.
	 */
	private scheduleUi(): void {
		if (this.uiFrame) return;
		this.uiFrame = window.requestAnimationFrame(() => {
			this.uiFrame = 0;
			this.ribbon?.update();
			this.selectionPane?.refresh();
			this.inspectorPane?.refresh();
			this.contextToolbar?.refresh();
			this.updateStatus();
		});
	}

	/** The editable text box of a shape on the current slide, if it has one. */
	private editableBoxOf(shapeId: string): HTMLElement | null {
		const slideEl = this.activeSlideEl;
		if (!slideEl) return null;
		for (const el of Array.from(slideEl.querySelectorAll<HTMLElement>("[data-shape-id]"))) {
			if (el.dataset.shapeId !== shapeId) continue;
			return el.querySelector<HTMLElement>('[data-editable="1"]');
		}
		return null;
	}

	/** Open the text editor on a shape, for Enter and F2. */
	private editTextOf(shapeId: string): void {
		const box = this.editableBoxOf(shapeId);
		if (box) this.editController?.beginAtEnd(box);
	}

	/**
	 * Typing on a selected shape opens its text and replaces what was there.
	 * Returns false for a shape with no text, leaving the key to its other uses.
	 */
	private replaceTextOf(shapeId: string, text: string): boolean {
		const box = this.editableBoxOf(shapeId);
		if (!box) return false;
		this.editController?.beginReplacing(box, text);
		return true;
	}

	/** A short description of the selection, shown in the status bar. */
	private updateStatus(): void {
		const ctx = this.context();
		if (!ctx || !this.viewer) return;
		const shapes = selectedShapes(ctx);
		if (shapes.length === 0) {
			this.viewer.setStatus("");
			return;
		}
		if (shapes.length > 1) {
			this.viewer.setStatus(t("status.multiple", { n: shapes.length }));
			return;
		}
		const shape = shapes[0];
		const f = shape.frame;
		this.viewer.setStatus(
			`${shape.name || shape.kind} · ${Math.round(f.w)} × ${Math.round(f.h)} · ` +
				`${Math.round(f.x)}, ${Math.round(f.y)}`,
		);
	}

	/** Right-clicking a thumbnail. */
	private showSlideMenu(index: number, event: MouseEvent): void {
		const ctx = this.context();
		if (!ctx) return;
		const menu = new Menu();
		menu.addItem((item) =>
			item
				.setTitle(t("cmd.newSlide"))
				.setIcon("file-plus")
				.onClick(() => this.runSlide((c) => newSlide(c))),
		);
		menu.addItem((item) =>
			item
				.setTitle(t("cmd.duplicateSlide"))
				.setIcon("files")
				.onClick(() => this.runSlide(duplicateCurrentSlide)),
		);
		menu.addSeparator();
		menu.addItem((item) =>
			item
				.setTitle(t("cmd.moveSlideUp"))
				.setIcon("arrow-up")
				.setDisabled(index === 0)
				.onClick(() => this.runSlide((c) => moveCurrentSlide(c, -1))),
		);
		menu.addItem((item) =>
			item
				.setTitle(t("cmd.moveSlideDown"))
				.setIcon("arrow-down")
				.setDisabled(index >= ctx.deck.slides.length - 1)
				.onClick(() => this.runSlide((c) => moveCurrentSlide(c, 1))),
		);
		menu.addSeparator();
		menu.addItem((item) =>
			item
				.setTitle(t("cmd.deleteSlide"))
				.setIcon("trash-2")
				.setDisabled(!canDeleteSlide(ctx))
				.onClick(() => this.run(deleteCurrentSlide)),
		);
		menu.showAtMouseEvent(event);
	}

	private context(): CommandContext | null {
		const loaded = this.loaded;
		const viewer = this.viewer;
		const editor = this.deckEditor;
		if (!loaded || !viewer || !editor) return null;
		const slide = loaded.deck.slides[viewer.currentSlideNumber - 1];
		if (!slide) return null;
		return { editor, pkg: loaded.pkg, deck: loaded.deck, slide, selection: this.selection };
	}

	private run(fn: (ctx: CommandContext) => unknown): void {
		const ctx = this.context();
		if (!ctx) return;
		try {
			fn(ctx);
		} catch (error) {
			new Notice(t("notice.editFailed", { message: (error as Error).message }), 8000);
		}
		this.scheduleUi();
	}

	/** Run a slide-level command and navigate to the slide it returns. */
	private runSlide(fn: (ctx: CommandContext) => number): void {
		const ctx = this.context();
		if (!ctx) return;
		let target = -1;
		try {
			target = fn(ctx);
		} catch (error) {
			new Notice(t("notice.editFailed", { message: (error as Error).message }), 8000);
			return;
		}
		if (target >= 0) this.viewer?.go(target);
		this.scheduleUi();
	}

	// ---------------------------------------------------------- ribbon host

	private ribbonHost(): RibbonHost {
		return {
			ctx: () => this.context(),
			run: (fn) => this.run(fn),
			runSlide: (fn) => this.runSlide(fn),
			canEdit: () => this.loaded !== null,
			zoomIn: () => this.viewer?.zoomIn(),
			zoomOut: () => this.viewer?.zoomOut(),
			zoomToFit: () => this.viewer?.zoomToFit(),
			toggleNotes: () => {
				this.viewer?.showNotes(!this.viewer.notesShown);
				this.scheduleUi();
			},
			notesShown: () => this.viewer?.notesShown ?? false,
			toggleThumbnails: () => {
				this.showThumbnails = !this.showThumbnails;
				this.rebuildViewer();
			},
			toggleSelectionPane: () => {
				this.showSelectionPane = !this.showSelectionPane;
				this.rebuildViewer();
			},
			selectionPaneShown: () => this.showSelectionPane,
			toggleRulers: () => {
				this.showRulers = !this.showRulers;
				this.rebuildViewer();
			},
			rulersShown: () => this.showRulers,
			addGuide: (orientation) => this.guides?.addCentre(orientation),
			toggleCrop: () => this.crop?.toggle(),
			cropActive: () => this.crop?.active ?? false,
			canCrop: () => this.crop?.canCrop() ?? false,
			resetCrop: () => this.crop?.reset(),
			clearGuides: () => this.guides?.clearAll(),
			hasGuides: () => (this.guides?.all.length ?? 0) > 0,
			save: () => void this.save(),
			isDirty: () => this.loaded?.pkg.isDirty ?? false,
			undo: () => this.undo(),
			redo: () => this.redo(),
			canUndo: () => this.deckEditor?.canUndo ?? false,
			canRedo: () => this.deckEditor?.canRedo ?? false,
			selectAll: () => {
				this.shapeEditor?.selectAll();
				this.scheduleUi();
			},
			pickImage: () => this.pickImage(),
			pickTable: () => this.pickTable(),
			pickLayout: () => this.pickLayout(),
			pickHyperlink: () => this.pickHyperlink(),
			findReplace: () => this.openFindReplace(),
			tableSelection: this.tableSelection,
			slideBackground: () => {
				const ctx = this.context();
				const background = ctx?.slide.background;
				return background?.kind === "solid" ? background.color : null;
			},
			exportPng: () => void this.exportCurrentSlide(),
			extractMarkdown: () => {
				if (this.loaded) void this.plugin.extractMarkdown(this.loaded, this.file);
			},
			openExternally: () => this.plugin.openExternally(this.file),
		};
	}

	/** Rebuild only the viewer, e.g. after toggling the thumbnail rail. */
	private rebuildViewer(): void {
		const current = this.viewer?.currentSlideNumber ?? 1;
		this.selectionPane?.destroy();
		this.selectionPane = null;
		this.sideSplitter?.destroy();
		this.sideSplitter = null;
		this.inspectorPane?.destroy();
		this.inspectorPane = null;
		this.viewer?.destroy();
		this.createViewer(current);
		this.scheduleUi();
	}

	// -------------------------------------------------------------- insert

	private pickImage(): void {
		new ImagePickerModal(this.app, (file) => {
			void (async () => {
				try {
					const bytes = new Uint8Array(await this.app.vault.readBinary(file));
					const size = await imageDimensions(bytes, `image/${file.extension}`);
					this.run((ctx) =>
						insertPicture(ctx, {
							bytes,
							extension: file.extension,
							name: file.name,
							width: size?.width,
							height: size?.height,
						}),
					);
				} catch (error) {
					new Notice(t("notice.imageFailed", { message: (error as Error).message }), 8000);
				}
			})();
		}).open();
	}

	/**
	 * Find and replace over the whole deck.
	 *
	 * The search reads every slide's XML, not just the one on screen, so a hit
	 * on slide 30 is found without rendering slides 2 to 29.
	 */
	private openFindReplace(): void {
		const loaded = this.loaded;
		if (!loaded) return;
		new FindReplaceModal(this.app, {
			search: (query, matchCase) => findMatches(loaded.deck, query, { matchCase }),
			replaceAll: (query, replacement, matchCase) => {
				const ctx = this.context();
				if (!ctx) return 0;
				const count = replaceAll(ctx, query, replacement, { matchCase }, t("modal.replaceAll"));
				if (count > 0) this.scheduleUi();
				return count;
			},
			reveal: (match) => {
				this.viewer?.go(match.slideIndex);
				this.selection.set(match.slideIndex, [match.shapeId]);
				this.viewer?.focus();
				this.scheduleUi();
			},
		}).open();
	}

	private pickHyperlink(): void {
		const ctx = this.context();
		if (!ctx) return;
		const current = hyperlinkState(ctx) ?? "";
		new PromptModal(this.app, t("modal.hyperlink"), current, (value) => {
			this.run((c) => setHyperlink(c, value === "" ? null : value));
		}).open();
	}

	private pickTable(): void {
		new TableSizeModal(this.app, (rows, columns) => {
			this.run((ctx) => insertTable(ctx, rows, columns));
		}).open();
	}

	private pickLayout(): void {
		const ctx = this.context();
		if (!ctx) return;
		const layouts = listLayouts(ctx);
		if (layouts.length === 0) {
			new Notice(t("notice.noLayouts"));
			return;
		}
		new LayoutPickerModal(this.app, layouts, (layout) => {
			this.runSlide((c) => newSlide(c, layout.path));
		}).open();
	}

	// -------------------------------------------------------- context menu

	private showContextMenu(event: MouseEvent): void {
		const ctx = this.context();
		if (!ctx) return;
		const menu = new Menu();
		const shapes = selectedShapes(ctx);
		const hasGroup = shapes.some((s) => s.kind === "group");

		if (shapes.length > 0) {
			menu.addItem((item) =>
				item.setTitle(t("cmd.cut")).setIcon("scissors").onClick(() => this.run(cutSelection)),
			);
			menu.addItem((item) =>
				item.setTitle(t("cmd.copy")).setIcon("copy").onClick(() => this.run(copySelection)),
			);
			menu.addItem((item) =>
				item
					.setTitle(t("cmd.duplicate"))
					.setIcon("copy-plus")
					.onClick(() => this.run(duplicateSelection)),
			);
			menu.addItem((item) =>
				item.setTitle(t("cmd.delete")).setIcon("trash").onClick(() => this.run(deleteSelection)),
			);
			menu.addSeparator();
			menu.addItem((item) =>
				item
					.setTitle(t("cmd.bringToFront"))
					.setIcon("bring-to-front")
					.onClick(() => this.run((c) => reorderSelection(c, "front"))),
			);
			menu.addItem((item) =>
				item
					.setTitle(t("cmd.sendToBack"))
					.setIcon("send-to-back")
					.onClick(() => this.run((c) => reorderSelection(c, "back"))),
			);
			if (shapes.length > 1 || hasGroup) menu.addSeparator();
			if (shapes.length > 1) {
				menu.addItem((item) =>
					item.setTitle(t("cmd.group")).setIcon("group").onClick(() => this.run(groupSelection)),
				);
			}
			if (hasGroup) {
				menu.addItem((item) =>
					item
						.setTitle(t("cmd.ungroup"))
						.setIcon("ungroup")
						.onClick(() => this.run(ungroupSelection)),
				);
			}
		} else {
			menu.addItem((item) =>
				item
					.setTitle(t("cmd.paste"))
					.setIcon("clipboard-paste")
					.setDisabled(!hasClipboard())
					.onClick(() => this.run((c) => pasteClipboard(c))),
			);
			menu.addItem((item) =>
				item
					.setTitle(t("cmd.newSlide"))
					.setIcon("file-plus")
					.onClick(() => this.runSlide((c) => newSlide(c))),
			);
			menu.addItem((item) =>
				item
					.setTitle(t("cmd.deleteSlide"))
					.setIcon("trash-2")
					.setDisabled(!canDeleteSlide(ctx))
					.onClick(() => this.run(deleteCurrentSlide)),
			);
		}
		menu.showAtMouseEvent(event);
	}

	// ----------------------------------------------------------- shortcuts

	/**
	 * An image pasted from outside Obsidian.
	 *
	 * The internal clipboard — shapes copied from a slide — is handled by
	 * Cmd/Ctrl+V, so this only takes over when the system clipboard is carrying
	 * a picture, and never while a caret is in a text box, where a paste
	 * belongs to the text.
	 */
	private onPaste = (event: ClipboardEvent): void => {
		if (this.editController?.isEditing || !this.loaded) return;
		const chosen = chooseImage(Array.from(event.clipboardData?.files ?? []));
		if (!chosen) return;
		event.preventDefault();
		event.stopPropagation();
		void this.insertClipboardImage(chosen.file, chosen.extension);
	};

	private async insertClipboardImage(file: File, extension: string): Promise<void> {
		try {
			const bytes = new Uint8Array(await file.arrayBuffer());
			const size = await imageDimensions(bytes, file.type);
			this.run((ctx) =>
				insertPicture(ctx, {
					bytes,
					extension,
					name: clipboardImageName(file.name, extension),
					width: size?.width,
					height: size?.height,
				}),
			);
		} catch (error) {
			new Notice(t("notice.imageFailed", { message: (error as Error).message }), 8000);
		}
	}

	private onKeyDown = (event: KeyboardEvent): void => {
		if (!(event.metaKey || event.ctrlKey)) return;
		const key = event.key.toLowerCase();

		// Saving works even mid-edit; everything else would fight the caret.
		if (key === "s") {
			event.preventDefault();
			event.stopPropagation();
			void this.save();
			return;
		}
		if (this.editController?.isEditing) return;

		const handlers: Record<string, () => void> = {
			z: () => (event.shiftKey ? this.redo() : this.undo()),
			y: () => this.redo(),
			c: () => this.run(copySelection),
			x: () => this.run(cutSelection),
			v: () => this.run((ctx) => pasteClipboard(ctx)),
			d: () => this.run(duplicateSelection),
			a: () => {
				this.shapeEditor?.selectAll();
				this.scheduleUi();
			},
			f: () => this.openFindReplace(),
		};
		const handler = handlers[key];
		if (!handler) return;
		event.preventDefault();
		event.stopPropagation();
		handler();
	};

	undo(): void {
		this.editController?.commit();
		this.deckEditor?.undo();
	}

	redo(): void {
		this.editController?.commit();
		this.deckEditor?.redo();
	}

	// ---------------------------------------------------------------- save

	async save(): Promise<void> {
		const file = this.file;
		const loaded = this.loaded;
		if (!file || !loaded) return;
		this.editController?.commit();
		if (!loaded.pkg.isDirty) {
			new Notice(t("notice.noChanges"));
			return;
		}
		try {
			const result = await saveDeck(this.app, file, loaded.pkg, this.baseMtime);
			this.baseMtime = file.stat.mtime;
			this.plugin.decks.touch(file.path, file.stat.mtime, file.stat.size);
			this.viewer?.setDirty(false);
			this.scheduleUi();
			new Notice(
				result.backupPath
					? t("notice.savedBackup", { path: result.backupPath })
					: t("notice.saved"),
			);
		} catch (error) {
			if (error instanceof ConflictError) new Notice(error.message, 8000);
			else new Notice(t("notice.saveFailed", { message: (error as Error).message }), 8000);
		}
	}

	async exportCurrentSlide(): Promise<void> {
		if (!this.loaded || !this.viewer) return;
		await this.plugin.exportSlidePng(this.loaded, this.file, this.viewer.currentSlideNumber);
	}

	get hasUnsavedChanges(): boolean {
		return this.loaded?.pkg.isDirty ?? false;
	}

	private renderError(file: TFile, error: Error): void {
		const box = this.contentEl.createDiv({ cls: "pptx-message pptx-error" });
		box.createDiv({ cls: "pptx-error-title", text: t("view.openFailed") });
		box.createDiv({ cls: "pptx-error-detail", text: error.message });
		const open = box.createEl("button", { text: t("view.openExternally") });
		open.addEventListener("click", () => this.plugin.openExternally(file));
	}

	private teardown(): void {
		if (this.uiFrame) window.cancelAnimationFrame(this.uiFrame);
		this.uiFrame = 0;
		this.contentEl.removeEventListener("keydown", this.onKeyDown, { capture: true });
		this.contentEl.removeEventListener("paste", this.onPaste, { capture: true });
		this.selectionPane?.destroy();
		this.selectionPane = null;
		this.sideSplitter?.destroy();
		this.sideSplitter = null;
		this.inspectorPane?.destroy();
		this.inspectorPane = null;
		this.contextToolbar?.destroy();
		this.contextToolbar = null;
		this.viewer?.destroy();
		this.ribbon?.destroy();
		this.viewer = null;
		this.ribbon = null;
		this.editController = null;
		this.shapeEditor = null;
		this.guides = null;
		this.crop = null;
		this.deckEditor = null;
		this.selection.clear();
		this.tableSelection.clear();
		this.activeSlideEl = null;
		this.loaded = null;
	}
}
