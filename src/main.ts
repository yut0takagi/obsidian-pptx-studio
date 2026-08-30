import { Notice, Plugin, TFile, normalizePath } from "obsidian";
import { DeckCache, type LoadedDeck } from "./DeckCache";
import { deckToMarkdown } from "./export/markdown";
import { pngFileName, renderSlideToPng } from "./export/png";
import { DEFAULT_SETTINGS, type PptxViewerSettings } from "./settings";
import { PptxSettingsTab } from "./view/SettingsTab";
import { PptxView, VIEW_TYPE_PPTX } from "./view/PptxView";
import { registerPptxEmbeds } from "./view/embed";

export default class PptxViewerPlugin extends Plugin {
	settings: PptxViewerSettings = DEFAULT_SETTINGS;
	decks!: DeckCache;

	async onload(): Promise<void> {
		await this.loadSettings();
		this.decks = new DeckCache(this.app);

		this.registerView(VIEW_TYPE_PPTX, (leaf) => new PptxView(leaf, this));

		// This is what makes .pptx visible in the file explorer and resolvable by
		// [[wiki links]]; without it Obsidian treats the extension as unknown.
		try {
			this.registerExtensions(["pptx"], VIEW_TYPE_PPTX);
		} catch {
			new Notice(
				"PPTX Viewer: another plugin already handles .pptx files. Disable it to use this one.",
				10000,
			);
		}

		registerPptxEmbeds(this);
		this.addSettingTab(new PptxSettingsTab(this));
		this.registerCommands();
		this.registerVaultEvents();
	}

	onunload(): void {
		this.decks.clear();
	}

	private registerCommands(): void {
		this.addCommand({
			id: "save-deck",
			name: "Save deck",
			checkCallback: (checking) => {
				const view = this.activeDeckView();
				if (!view?.hasUnsavedChanges) return false;
				if (!checking) void view.save();
				return true;
			},
		});

		this.addCommand({
			id: "export-slide-png",
			name: "Export current slide as PNG",
			checkCallback: (checking) => {
				const view = this.activeDeckView();
				if (!view?.file) return false;
				if (!checking) void view.exportCurrentSlide();
				return true;
			},
		});

		this.addCommand({
			id: "extract-markdown",
			name: "Extract deck text to a Markdown note",
			checkCallback: (checking) => {
				const file = this.targetDeckFile();
				if (!file) return false;
				if (!checking) {
					void this.decks
						.get(file)
						.then((loaded) => this.extractMarkdown(loaded, file))
						.catch((error) => new Notice(`Could not read the deck: ${error.message}`));
				}
				return true;
			},
		});

		this.addCommand({
			id: "open-externally",
			name: "Open deck in the default app",
			checkCallback: (checking) => {
				const file = this.targetDeckFile();
				if (!file) return false;
				if (!checking) this.openExternally(file);
				return true;
			},
		});
	}

	private registerVaultEvents(): void {
		this.registerEvent(
			this.app.vault.on("rename", (file, oldPath) => {
				if (file instanceof TFile && file.extension === "pptx") {
					this.decks.rename(oldPath, file.path);
				}
			}),
		);
		this.registerEvent(
			this.app.vault.on("delete", (file) => {
				if (file instanceof TFile && file.extension === "pptx") {
					this.decks.evict(file.path, true);
				}
			}),
		);
		this.registerEvent(
			this.app.vault.on("modify", (file) => {
				// A deck edited outside Obsidian must be re-read; one edited in here
				// is protected by the cache and keeps its unsaved changes.
				if (file instanceof TFile && file.extension === "pptx") {
					this.decks.evict(file.path);
				}
			}),
		);
	}

	private activeDeckView(): PptxView | null {
		const view = this.app.workspace.getActiveViewOfType(PptxView);
		return view ?? null;
	}

	/** The deck to act on: the open deck view, or a .pptx selected in the explorer. */
	private targetDeckFile(): TFile | null {
		const view = this.activeDeckView();
		if (view?.file) return view.file;
		const active = this.app.workspace.getActiveFile();
		return active?.extension === "pptx" ? active : null;
	}

	// ------------------------------------------------------------- actions

	async exportSlidePng(
		loaded: LoadedDeck,
		file: TFile | null,
		slideNumber: number,
	): Promise<void> {
		const slide = loaded.deck.slides[slideNumber - 1];
		if (!slide) return;
		try {
			const png = await renderSlideToPng(
				loaded.deck,
				slide,
				loaded.pkg,
				this.settings.exportScale,
			);
			const name = pngFileName(file?.basename ?? loaded.deck.title, slideNumber);
			const path = await this.resolveExportPath(name, file?.path ?? "");
			await this.app.vault.createBinary(path, png);
			new Notice(`Exported ${path}`);
		} catch (error) {
			new Notice(`Could not export the slide: ${(error as Error).message}`, 8000);
		}
	}

	async extractMarkdown(loaded: LoadedDeck, file: TFile | null): Promise<void> {
		const markdown = deckToMarkdown(loaded.deck, {
			includeNotes: this.settings.includeNotes,
			includeSourceLink: this.settings.includeSourceLink,
			sourcePath: file?.path ?? "",
		});
		const folder = file?.parent?.path ?? "";
		const base = `${file?.basename ?? loaded.deck.title} (slides)`;
		const path = await this.availablePath(folder, base, "md");
		try {
			const created = await this.app.vault.create(path, markdown);
			await this.app.workspace.getLeaf("tab").openFile(created);
		} catch (error) {
			new Notice(`Could not create the note: ${(error as Error).message}`, 8000);
		}
	}

	openExternally(file: TFile | null): void {
		if (!file) return;
		const opener = this.app as unknown as { openWithDefaultApp?: (path: string) => unknown };
		if (typeof opener.openWithDefaultApp === "function") {
			opener.openWithDefaultApp(file.path);
			return;
		}
		new Notice("This build of Obsidian cannot open files in an external app.");
	}

	// --------------------------------------------------------------- paths

	private async resolveExportPath(name: string, sourcePath: string): Promise<string> {
		if (this.settings.exportFolder) {
			const folder = normalizePath(this.settings.exportFolder);
			if (!this.app.vault.getAbstractFileByPath(folder)) {
				await this.app.vault.createFolder(folder).catch(() => undefined);
			}
			return this.uniquePath(`${folder}/${name}`);
		}
		const manager = this.app.fileManager as unknown as {
			getAvailablePathForAttachment?: (name: string, sourcePath?: string) => Promise<string>;
		};
		if (typeof manager.getAvailablePathForAttachment === "function") {
			return manager.getAvailablePathForAttachment(name, sourcePath);
		}
		return this.uniquePath(name);
	}

	private async availablePath(folder: string, base: string, extension: string): Promise<string> {
		const prefix = folder ? `${folder}/` : "";
		return this.uniquePath(`${prefix}${base}.${extension}`);
	}

	private uniquePath(path: string): string {
		const normalised = normalizePath(path);
		if (!this.app.vault.getAbstractFileByPath(normalised)) return normalised;
		const dot = normalised.lastIndexOf(".");
		const stem = dot === -1 ? normalised : normalised.slice(0, dot);
		const ext = dot === -1 ? "" : normalised.slice(dot);
		for (let i = 1; i < 1000; i++) {
			const candidate = `${stem} ${i}${ext}`;
			if (!this.app.vault.getAbstractFileByPath(candidate)) return candidate;
		}
		return `${stem} ${Date.now()}${ext}`;
	}

	// ------------------------------------------------------------ settings

	async loadSettings(): Promise<void> {
		this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
	}

	async saveSettings(): Promise<void> {
		await this.saveData(this.settings);
	}
}
