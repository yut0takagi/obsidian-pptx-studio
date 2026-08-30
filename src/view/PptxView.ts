import { FileView, Menu, Notice, type TFile, type WorkspaceLeaf } from "obsidian";
import type { LoadedDeck } from "../DeckCache";
import { DeckEditor } from "../edit/DeckEditor";
import { EditController } from "../edit/EditController";
import { Selection } from "../edit/Selection";
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
	listLayouts,
	newSlide,
	reorderSlide,
} from "../edit/slideCommands";
import { rebuildDeck } from "../pptx/parse";
import { DeckViewer } from "../render/DeckViewer";
import { Ribbon } from "../ui/Ribbon";
import {
	ImagePickerModal,
	LayoutPickerModal,
	PromptModal,
	TableSizeModal,
	imageDimensions,
} from "../ui/modals";
import { type RibbonHost, buildTabs } from "../ui/tabs";
import type PptxViewerPlugin from "../main";

export const VIEW_TYPE_PPTX = "pptx-viewer";

/** The full-tab deck editor: ribbon, slide canvas, selection, save. */
export class PptxView extends FileView {
	private viewer: DeckViewer | null = null;
	private ribbon: Ribbon | null = null;
	private editController: EditController | null = null;
	private shapeEditor: ShapeEditor | null = null;
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
		private readonly plugin: PptxViewerPlugin,
	) {
		super(leaf);
		this.selection.onChange(() => {
			this.tableSelection.retain(this.selection.ids);
			this.ribbon?.update();
		});
		this.tableSelection.onChange(() => {
			this.tableSelection.paint(this.activeSlideEl);
			this.ribbon?.update();
		});
	}

	getViewType(): string {
		return VIEW_TYPE_PPTX;
	}

	getIcon(): string {
		return "presentation";
	}

	getDisplayText(): string {
		return this.file?.basename ?? "Presentation";
	}

	canAcceptExtension(extension: string): boolean {
		return extension === "pptx";
	}

	async onLoadFile(file: TFile): Promise<void> {
		this.teardown();
		this.contentEl.empty();
		this.contentEl.addClass("pptx-view");
		this.showThumbnails = this.plugin.settings.showThumbnails;

		const loading = this.contentEl.createDiv({ cls: "pptx-message", text: "Reading deck…" });
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
			onChanged: (rebuild) => this.afterEdit(rebuild),
		});

		this.editController = new EditController({
			isEnabled: () => true,
			editor: this.deckEditor,
			onCancelled: () => this.viewer?.invalidate(),
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

		this.shapeEditor = new ShapeEditor({
			selection: this.selection,
			editor: this.deckEditor,
			// Dragging is off while a text box is open, so a stray drag inside the
			// caret's box cannot move the shape out from under the editor.
			isEnabled: () => !this.editController?.isEditing,
			getScale: () => this.viewer?.currentScale() ?? 1,
			getContext: () => this.context(),
			onContextMenu: (event) => this.showContextMenu(event),
			onCellPointerDown: (shape, row, column, additive) =>
				this.tableSelection.select(shape.id, row, column, additive),
		});

		this.ribbon = new Ribbon(this.contentEl, buildTabs(this.ribbonHost()));
		this.createViewer(1);
		this.contentEl.addEventListener("keydown", this.onKeyDown, { capture: true });
		this.viewer?.focus();
		this.ribbon.update();
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
			onRendered: (slideEl) => {
				this.activeSlideEl = slideEl;
				this.tableSelection.paint(slideEl);
			},
			onReorder: (from, to) => {
				this.run((ctx) => reorderSlide(ctx, from, to));
				this.viewer?.go(to);
			},
		});
		if (slideNumber > 1) this.viewer.go(slideNumber - 1);
		this.viewer.setDirty(loaded.pkg.isDirty);
	}

	/** Rebuild the model and repaint after an edit. */
	private afterEdit(rebuild: boolean): void {
		const loaded = this.loaded;
		if (!loaded) return;
		if (rebuild) {
			loaded.deck = rebuildDeck(loaded.pkg, this.file?.basename ?? "Deck");
			this.viewer?.setDeck(loaded.deck);
			// Shapes can vanish from under a selection: deleted, ungrouped, undone.
			const slide = loaded.deck.slides[this.selection.slideIndex];
			const existing = new Set((slide?.shapes ?? []).filter((s) => s.source).map((s) => s.id));
			this.selection.retain(existing);
			this.tableSelection.retain(this.selection.ids);
			this.shapeEditor?.refresh();
			this.tableSelection.paint(this.activeSlideEl);
		}
		this.viewer?.setDirty(true);
		this.ribbon?.update();
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
			new Notice(`That edit failed: ${(error as Error).message}`, 8000);
		}
		this.ribbon?.update();
	}

	/** Run a slide-level command and navigate to the slide it returns. */
	private runSlide(fn: (ctx: CommandContext) => number): void {
		const ctx = this.context();
		if (!ctx) return;
		let target = -1;
		try {
			target = fn(ctx);
		} catch (error) {
			new Notice(`That edit failed: ${(error as Error).message}`, 8000);
			return;
		}
		if (target >= 0) this.viewer?.go(target);
		this.ribbon?.update();
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
				this.ribbon?.update();
			},
			notesShown: () => this.viewer?.notesShown ?? false,
			toggleThumbnails: () => {
				this.showThumbnails = !this.showThumbnails;
				this.rebuildViewer();
			},
			save: () => void this.save(),
			isDirty: () => this.loaded?.pkg.isDirty ?? false,
			undo: () => this.undo(),
			redo: () => this.redo(),
			canUndo: () => this.deckEditor?.canUndo ?? false,
			canRedo: () => this.deckEditor?.canRedo ?? false,
			selectAll: () => {
				this.shapeEditor?.selectAll();
				this.ribbon?.update();
			},
			pickImage: () => this.pickImage(),
			pickTable: () => this.pickTable(),
			pickLayout: () => this.pickLayout(),
			pickHyperlink: () => this.pickHyperlink(),
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
		this.viewer?.destroy();
		this.createViewer(current);
		this.ribbon?.update();
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
					new Notice(`Could not insert the image: ${(error as Error).message}`, 8000);
				}
			})();
		}).open();
	}

	private pickHyperlink(): void {
		const ctx = this.context();
		if (!ctx) return;
		const current = hyperlinkState(ctx) ?? "";
		new PromptModal(this.app, "Hyperlink", current, (value) => {
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
			new Notice("This deck has no slide layouts to choose from.");
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
				item.setTitle("Cut").setIcon("scissors").onClick(() => this.run(cutSelection)),
			);
			menu.addItem((item) =>
				item.setTitle("Copy").setIcon("copy").onClick(() => this.run(copySelection)),
			);
			menu.addItem((item) =>
				item.setTitle("Duplicate").setIcon("copy-plus").onClick(() => this.run(duplicateSelection)),
			);
			menu.addItem((item) =>
				item.setTitle("Delete").setIcon("trash").onClick(() => this.run(deleteSelection)),
			);
			menu.addSeparator();
			menu.addItem((item) =>
				item
					.setTitle("Bring to front")
					.setIcon("bring-to-front")
					.onClick(() => this.run((c) => reorderSelection(c, "front"))),
			);
			menu.addItem((item) =>
				item
					.setTitle("Send to back")
					.setIcon("send-to-back")
					.onClick(() => this.run((c) => reorderSelection(c, "back"))),
			);
			if (shapes.length > 1 || hasGroup) menu.addSeparator();
			if (shapes.length > 1) {
				menu.addItem((item) =>
					item.setTitle("Group").setIcon("group").onClick(() => this.run(groupSelection)),
				);
			}
			if (hasGroup) {
				menu.addItem((item) =>
					item.setTitle("Ungroup").setIcon("ungroup").onClick(() => this.run(ungroupSelection)),
				);
			}
		} else {
			menu.addItem((item) =>
				item
					.setTitle("Paste")
					.setIcon("clipboard-paste")
					.setDisabled(!hasClipboard())
					.onClick(() => this.run((c) => pasteClipboard(c))),
			);
			menu.addItem((item) =>
				item
					.setTitle("New slide")
					.setIcon("file-plus")
					.onClick(() => this.runSlide((c) => newSlide(c))),
			);
			menu.addItem((item) =>
				item
					.setTitle("Delete slide")
					.setIcon("trash-2")
					.setDisabled(!canDeleteSlide(ctx))
					.onClick(() => this.run(deleteCurrentSlide)),
			);
		}
		menu.showAtMouseEvent(event);
	}

	// ----------------------------------------------------------- shortcuts

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
				this.ribbon?.update();
			},
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
			new Notice("No changes to save.");
			return;
		}
		try {
			const result = await saveDeck(this.app, file, loaded.pkg, this.baseMtime);
			this.baseMtime = file.stat.mtime;
			this.plugin.decks.touch(file.path, file.stat.mtime, file.stat.size);
			this.viewer?.setDirty(false);
			this.ribbon?.update();
			new Notice(
				result.backupPath
					? `Saved. A backup of the original is at ${result.backupPath}.`
					: "Saved.",
			);
		} catch (error) {
			if (error instanceof ConflictError) new Notice(error.message, 8000);
			else new Notice(`Could not save: ${(error as Error).message}`, 8000);
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
		box.createDiv({ cls: "pptx-error-title", text: "This deck could not be opened" });
		box.createDiv({ cls: "pptx-error-detail", text: error.message });
		const open = box.createEl("button", { text: "Open in the default app" });
		open.addEventListener("click", () => this.plugin.openExternally(file));
	}

	private teardown(): void {
		this.contentEl.removeEventListener("keydown", this.onKeyDown, { capture: true });
		this.viewer?.destroy();
		this.ribbon?.destroy();
		this.viewer = null;
		this.ribbon = null;
		this.editController = null;
		this.shapeEditor = null;
		this.deckEditor = null;
		this.selection.clear();
		this.tableSelection.clear();
		this.activeSlideEl = null;
		this.loaded = null;
	}
}
