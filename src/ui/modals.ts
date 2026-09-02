import {
	App,
	FuzzySuggestModal,
	Modal,
	Notice,
	Setting,
	SuggestModal,
	type TFile,
} from "obsidian";
import type { TextMatch } from "../edit/findReplace";
import { t } from "../i18n";

const IMAGE_EXTENSIONS = new Set(["png", "jpg", "jpeg", "gif", "bmp", "svg", "webp"]);

/** Pick an image from the vault to insert onto a slide. */
export class ImagePickerModal extends FuzzySuggestModal<TFile> {
	constructor(
		app: App,
		private readonly onPick: (file: TFile) => void,
	) {
		super(app);
		this.setPlaceholder(t("modal.pickImage"));
	}

	getItems(): TFile[] {
		return this.app.vault
			.getFiles()
			.filter((file) => IMAGE_EXTENSIONS.has(file.extension.toLowerCase()))
			.sort((a, b) => b.stat.mtime - a.stat.mtime);
	}

	getItemText(file: TFile): string {
		return file.path;
	}

	onChooseItem(file: TFile): void {
		this.onPick(file);
	}
}

/** Pick the layout a new slide is built on. */
export class LayoutPickerModal extends SuggestModal<{ path: string; name: string }> {
	constructor(
		app: App,
		private readonly layouts: { path: string; name: string }[],
		private readonly onPick: (layout: { path: string; name: string }) => void,
	) {
		super(app);
		this.setPlaceholder(t("modal.pickLayout"));
	}

	getSuggestions(query: string): { path: string; name: string }[] {
		const needle = query.toLowerCase();
		return this.layouts.filter((layout) => layout.name.toLowerCase().includes(needle));
	}

	renderSuggestion(layout: { path: string; name: string }, el: HTMLElement): void {
		el.createDiv({ text: layout.name });
		el.createDiv({ cls: "pptx-suggest-note", text: layout.path });
	}

	onChooseSuggestion(layout: { path: string; name: string }): void {
		this.onPick(layout);
	}
}

/** Ask for a table's dimensions before inserting it. */
export class TableSizeModal extends Modal {
	private rows = 3;
	private columns = 3;

	constructor(
		app: App,
		private readonly onSubmit: (rows: number, columns: number) => void,
	) {
		super(app);
	}

	onOpen(): void {
		this.titleEl.setText(t("modal.tableTitle"));
		const { contentEl } = this;

		new Setting(contentEl).setName(t("modal.rows")).addSlider((slider) =>
			slider
				.setLimits(1, 12, 1)
				.setValue(this.rows)
				.onChange((value) => {
					this.rows = value;
				}),
		);
		new Setting(contentEl).setName(t("modal.columns")).addSlider((slider) =>
			slider
				.setLimits(1, 10, 1)
				.setValue(this.columns)
				.onChange((value) => {
					this.columns = value;
				}),
		);

		new Setting(contentEl).addButton((button) =>
			button
				.setButtonText(t("modal.insert"))
				.setCta()
				.onClick(() => {
					this.close();
					this.onSubmit(this.rows, this.columns);
				}),
		);
	}

	onClose(): void {
		this.contentEl.empty();
	}
}

/** Ask for one line of text, e.g. a URL. */
export class PromptModal extends Modal {
	private value: string;

	constructor(
		app: App,
		private readonly title: string,
		initial: string,
		private readonly onSubmit: (value: string) => void,
	) {
		super(app);
		this.value = initial;
	}

	onOpen(): void {
		this.titleEl.setText(this.title);
		const setting = new Setting(this.contentEl).addText((text) =>
			text
				.setValue(this.value)
				.setPlaceholder("https://")
				.onChange((value) => {
					this.value = value;
				}),
		);
		const input = setting.controlEl.querySelector<HTMLInputElement>("input");
		input?.addEventListener("keydown", (event) => {
			if (event.key === "Enter") this.submit();
		});
		window.setTimeout(() => input?.focus(), 0);

		new Setting(this.contentEl).addButton((button) =>
			button.setButtonText(t("common.ok")).setCta().onClick(() => this.submit()),
		);
	}

	private submit(): void {
		this.close();
		this.onSubmit(this.value.trim());
	}

	onClose(): void {
		this.contentEl.empty();
	}
}

export interface FindReplaceHost {
	search: (query: string, matchCase: boolean) => TextMatch[];
	/** Returns how many were replaced, which can be fewer than were found. */
	replaceAll: (query: string, replacement: string, matchCase: boolean) => number;
	/** Go to a hit: show its slide and select the shape holding it. */
	reveal: (match: TextMatch) => void;
}

/** How much of a paragraph to show around a hit. */
const CONTEXT = 40;

/**
 * Find and replace across the deck.
 *
 * The results are the point: a replace-all over a deck you cannot see is a
 * leap of faith, so every hit is listed with the words either side of it and
 * the slide it sits on, and clicking one goes there.
 */
export class FindReplaceModal extends Modal {
	private query = "";
	private replacement = "";
	private matchCase = false;
	private resultsEl: HTMLElement | null = null;

	constructor(
		app: App,
		private readonly host: FindReplaceHost,
	) {
		super(app);
	}

	onOpen(): void {
		this.titleEl.setText(t("modal.findTitle"));
		const { contentEl } = this;

		const findSetting = new Setting(contentEl).setName(t("modal.find")).addText((text) =>
			text.setValue(this.query).onChange((value) => {
				this.query = value;
			}),
		);
		new Setting(contentEl).setName(t("modal.replaceWith")).addText((text) =>
			text.setValue(this.replacement).onChange((value) => {
				this.replacement = value;
			}),
		);
		new Setting(contentEl).setName(t("modal.matchCase")).addToggle((toggle) =>
			toggle.setValue(this.matchCase).onChange((value) => {
				this.matchCase = value;
				if (this.query) this.runSearch();
			}),
		);

		new Setting(contentEl)
			.addButton((button) => button.setButtonText(t("modal.findAll")).onClick(() => this.runSearch()))
			.addButton((button) =>
				button
					.setButtonText(t("modal.replaceAll"))
					.setCta()
					.onClick(() => this.runReplace()),
			);

		this.resultsEl = contentEl.createDiv({ cls: "pptx-find-results" });

		const input = findSetting.controlEl.querySelector<HTMLInputElement>("input");
		input?.addEventListener("keydown", (event) => {
			if (event.key === "Enter") this.runSearch();
		});
		window.setTimeout(() => input?.focus(), 0);
	}

	private runSearch(): void {
		const results = this.resultsEl;
		if (!results) return;
		results.empty();
		if (this.query === "") return;

		const matches = this.host.search(this.query, this.matchCase);
		results.createDiv({
			cls: "pptx-find-count",
			text: matches.length === 0 ? t("find.none") : t("find.count", { n: matches.length }),
		});
		for (const match of matches) {
			const row = results.createDiv({ cls: "pptx-find-hit" });
			row.createDiv({
				cls: "pptx-find-where",
				text: `${t("view.slideLabel", { n: match.slideIndex + 1 })} · ${match.shapeName}`,
			});
			const line = row.createDiv({ cls: "pptx-find-line" });
			const from = Math.max(0, match.start - CONTEXT);
			const to = Math.min(match.text.length, match.start + match.length + CONTEXT);
			if (from > 0) line.createSpan({ text: "…" });
			line.createSpan({ text: match.text.slice(from, match.start) });
			line.createEl("mark", { text: match.text.slice(match.start, match.start + match.length) });
			line.createSpan({ text: match.text.slice(match.start + match.length, to) });
			if (to < match.text.length) line.createSpan({ text: "…" });
			row.addEventListener("click", () => {
				this.close();
				this.host.reveal(match);
			});
		}
	}

	private runReplace(): void {
		if (this.query === "") return;
		const found = this.host.search(this.query, this.matchCase).length;
		const replaced = this.host.replaceAll(this.query, this.replacement, this.matchCase);
		this.close();
		new Notice(replaced === 0 ? t("find.none") : t("find.replaced", { n: replaced }));
		// A hit that straddles a line break or a field is reported but not
		// rewritten, and saying so beats leaving the count unexplained.
		if (found > replaced) new Notice(t("find.skipped", { n: found - replaced }));
	}

	onClose(): void {
		this.contentEl.empty();
	}
}

/** Natural pixel size of an image, so an inserted picture keeps its proportions. */
export async function imageDimensions(
	bytes: Uint8Array,
	mime: string,
): Promise<{ width: number; height: number } | null> {
	const url = URL.createObjectURL(new Blob([bytes.slice()], { type: mime }));
	try {
		return await new Promise((resolve) => {
			const img = new Image();
			img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
			img.onerror = () => resolve(null);
			img.src = url;
		});
	} finally {
		URL.revokeObjectURL(url);
	}
}
