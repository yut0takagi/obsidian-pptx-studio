import { FileView, Notice, type TFile, type WorkspaceLeaf } from "obsidian";
import { EditController } from "../edit/EditController";
import { ConflictError, saveDeck } from "../edit/save";
import { DeckViewer } from "../render/DeckViewer";
import { rebuildDeck } from "../pptx/parse";
import type { LoadedDeck } from "../DeckCache";
import type PptxViewerPlugin from "../main";

export const VIEW_TYPE_PPTX = "pptx-viewer";

/** The full-tab deck view: browse, edit text, save, export. */
export class PptxView extends FileView {
	private viewer: DeckViewer | null = null;
	private editor: EditController | null = null;
	private loaded: LoadedDeck | null = null;
	/** mtime the open deck was read from, used to detect outside edits. */
	private baseMtime = 0;

	constructor(
		leaf: WorkspaceLeaf,
		private readonly plugin: PptxViewerPlugin,
	) {
		super(leaf);
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

		const loading = this.contentEl.createDiv({ cls: "pptx-message", text: "Reading deck…" });
		try {
			const loaded = await this.plugin.decks.get(file);
			loading.remove();
			this.loaded = loaded;
			this.baseMtime = file.stat.mtime;
			this.build(loaded);
		} catch (error) {
			loading.remove();
			this.renderError(file, error as Error);
		}
	}

	async onUnloadFile(_file: TFile): Promise<void> {
		// Commit any in-progress edit so switching files never drops a keystroke.
		this.editor?.commit();
		this.teardown();
		this.contentEl.empty();
	}

	onunload(): void {
		this.editor?.commit();
		this.teardown();
	}

	private build(loaded: LoadedDeck): void {
		const settings = this.plugin.settings;

		this.editor = new EditController({
			isEnabled: () => true,
			onFinish: (changedPart) => {
				if (changedPart && this.loaded) {
					this.loaded.pkg.markDirty(changedPart);
					// Re-derive the model so the view reflects exactly what was written.
					this.loaded.deck = rebuildDeck(this.loaded.pkg, this.file?.basename ?? "Deck");
					this.viewer?.setDeck(this.loaded.deck);
					this.viewer?.setDirty(true);
				} else {
					this.viewer?.invalidate();
				}
			},
		});

		this.viewer = new DeckViewer(this.contentEl, {
			deck: loaded.deck,
			pkg: loaded.pkg,
			compact: false,
			controls: true,
			showThumbnails: settings.showThumbnails,
			showNotes: settings.showNotes,
			fitMode: settings.fitMode,
			editor: this.editor,
			onSave: () => void this.save(),
			onExportPng: (slideIndex) => void this.plugin.exportSlidePng(loaded, this.file, slideIndex),
			onExtractMarkdown: () => void this.plugin.extractMarkdown(loaded, this.file),
			onOpenExternal: () => this.plugin.openExternally(this.file),
		});
		this.viewer.setDirty(loaded.pkg.isDirty);

		this.contentEl.addEventListener("keydown", this.onKeyDown, { capture: true });
		this.viewer.focus();
	}

	/** Cmd/Ctrl+S saves while the deck view has focus. */
	private onKeyDown = (event: KeyboardEvent): void => {
		if (event.key !== "s" || !(event.metaKey || event.ctrlKey)) return;
		event.preventDefault();
		event.stopPropagation();
		void this.save();
	};

	async save(): Promise<void> {
		const file = this.file;
		const loaded = this.loaded;
		if (!file || !loaded) return;
		this.editor?.commit();
		if (!loaded.pkg.isDirty) {
			new Notice("No changes to save.");
			return;
		}
		try {
			const result = await saveDeck(this.app, file, loaded.pkg, this.baseMtime);
			this.baseMtime = file.stat.mtime;
			this.plugin.decks.touch(file.path, file.stat.mtime, file.stat.size);
			this.viewer?.setDirty(false);
			new Notice(
				result.backupPath
					? `Saved. A backup of the original is at ${result.backupPath}.`
					: "Saved.",
			);
		} catch (error) {
			if (error instanceof ConflictError) {
				new Notice(error.message, 8000);
			} else {
				new Notice(`Could not save: ${(error as Error).message}`, 8000);
			}
		}
	}

	/** Export the slide currently on screen. Used by the command palette. */
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
		this.viewer = null;
		this.editor = null;
		this.loaded = null;
	}
}
