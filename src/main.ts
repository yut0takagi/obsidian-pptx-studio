import { Notice, Plugin, TFile, normalizePath } from "obsidian";
import { setLanguage, t } from "./i18n";
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
		setLanguage(this.settings.language);
		this.decks = new DeckCache(this.app);

		this.registerView(VIEW_TYPE_PPTX, (leaf) => new PptxView(leaf, this));

		// This is what makes .pptx visible in the file explorer and resolvable by
		// [[wiki links]]; without it Obsidian treats the extension as unknown.
		try {
			this.registerExtensions(["pptx"], VIEW_TYPE_PPTX);
		} catch {
			new Notice(t("notice.extensionTaken"), 10000);
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
			name: t("palette.save"),
			checkCallback: (checking) => {
				const view = this.activeDeckView();
				if (!view?.hasUnsavedChanges) return false;
				if (!checking) void view.save();
				return true;
			},
		});

		this.addCommand({
			id: "export-slide-png",
			name: t("palette.exportPng"),
			checkCallback: (checking) => {
				const view = this.activeDeckView();
				if (!view?.file) return false;
				if (!checking) void view.exportCurrentSlide();
				return true;
			},
		});

		this.addCommand({
			id: "extract-markdown",
			name: t("palette.extract"),
			checkCallback: (checking) => {
				const file = this.targetDeckFile();
				if (!file) return false;
				if (!checking) {
					void this.decks
						.get(file)
						.then((loaded) => this.extractMarkdown(loaded, file))
						.catch(
							(error) => new Notice(t("notice.readFailed", { message: error.message })),
						);
				}
				return true;
			},
		});

		this.addCommand({
			id: "open-externally",
			name: t("palette.openExternally"),
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
			new Notice(t("notice.exported", { path }));
		} catch (error) {
			new Notice(t("notice.exportFailed", { message: (error as Error).message }), 8000);
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
			new Notice(t("notice.noteFailed", { message: (error as Error).message }), 8000);
		}
	}

	openExternally(file: TFile | null): void {
		if (!file) return;
		const opener = this.app as unknown as { openWithDefaultApp?: (path: string) => unknown };
		if (typeof opener.openWithDefaultApp === "function") {
			opener.openWithDefaultApp(file.path);
			return;
		}
		new Notice(t("notice.noExternalApp"));
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
		setLanguage(this.settings.language);
		await this.saveData(this.settings);
	}

	/**
	 * Re-open every deck view. The ribbon's labels are built once, so a language
	 * change has to rebuild them rather than waiting for the next file open.
	 */
	refreshViews(): void {
		for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_PPTX)) {
			const view = leaf.view;
			if (view instanceof PptxView && view.file) void view.onLoadFile(view.file);
		}
	}
}
