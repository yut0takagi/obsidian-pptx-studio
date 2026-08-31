import { App, FuzzySuggestModal, Modal, Setting, SuggestModal, type TFile } from "obsidian";
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
