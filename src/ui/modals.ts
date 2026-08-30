import { App, FuzzySuggestModal, Modal, Setting, SuggestModal, type TFile } from "obsidian";

const IMAGE_EXTENSIONS = new Set(["png", "jpg", "jpeg", "gif", "bmp", "svg", "webp"]);

/** Pick an image from the vault to insert onto a slide. */
export class ImagePickerModal extends FuzzySuggestModal<TFile> {
	constructor(
		app: App,
		private readonly onPick: (file: TFile) => void,
	) {
		super(app);
		this.setPlaceholder("Pick an image to insert");
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
		this.setPlaceholder("Pick a layout for the new slide");
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
		this.titleEl.setText("Insert table");
		const { contentEl } = this;

		new Setting(contentEl).setName("Rows").addSlider((slider) =>
			slider
				.setLimits(1, 12, 1)
				.setValue(this.rows)
				.setDynamicTooltip()
				.onChange((value) => {
					this.rows = value;
				}),
		);
		new Setting(contentEl).setName("Columns").addSlider((slider) =>
			slider
				.setLimits(1, 10, 1)
				.setValue(this.columns)
				.setDynamicTooltip()
				.onChange((value) => {
					this.columns = value;
				}),
		);

		new Setting(contentEl).addButton((button) =>
			button
				.setButtonText("Insert")
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
