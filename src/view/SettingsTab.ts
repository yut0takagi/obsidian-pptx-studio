import { PluginSettingTab, Setting } from "obsidian";
import type PptxViewerPlugin from "../main";

export class PptxSettingsTab extends PluginSettingTab {
	constructor(private readonly plugin: PptxViewerPlugin) {
		super(plugin.app, plugin);
	}

	display(): void {
		const { containerEl } = this;
		containerEl.empty();

		new Setting(containerEl).setName("Viewer").setHeading();

		new Setting(containerEl)
			.setName("Slide thumbnails")
			.setDesc("Show a thumbnail rail beside the slide.")
			.addToggle((toggle) =>
				toggle.setValue(this.plugin.settings.showThumbnails).onChange(async (value) => {
					this.plugin.settings.showThumbnails = value;
					await this.plugin.saveSettings();
				}),
			);

		new Setting(containerEl)
			.setName("Speaker notes")
			.setDesc("Show the notes pane when a deck opens. Toggle it any time with N.")
			.addToggle((toggle) =>
				toggle.setValue(this.plugin.settings.showNotes).onChange(async (value) => {
					this.plugin.settings.showNotes = value;
					await this.plugin.saveSettings();
				}),
			);

		new Setting(containerEl)
			.setName("Default zoom")
			.setDesc("How a slide is sized when the view opens.")
			.addDropdown((dropdown) =>
				dropdown
					.addOption("page", "Fit whole slide")
					.addOption("width", "Fit width")
					.setValue(this.plugin.settings.fitMode)
					.onChange(async (value) => {
						this.plugin.settings.fitMode = value === "width" ? "width" : "page";
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl).setName("Embeds in notes").setHeading();

		new Setting(containerEl)
			.setName("Embed height")
			.setDesc("Height in pixels of a deck embedded in a note.")
			.addSlider((slider) =>
				slider
					.setLimits(200, 900, 20)
					.setValue(this.plugin.settings.embedMaxHeight)
					.setDynamicTooltip()
					.onChange(async (value) => {
						this.plugin.settings.embedMaxHeight = value;
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl)
			.setName("Embed controls")
			.setDesc("Show page controls on embeds that are not pinned to one slide.")
			.addToggle((toggle) =>
				toggle.setValue(this.plugin.settings.embedControls).onChange(async (value) => {
					this.plugin.settings.embedControls = value;
					await this.plugin.saveSettings();
				}),
			);

		new Setting(containerEl).setName("Export").setHeading();

		new Setting(containerEl)
			.setName("PNG resolution")
			.setDesc("Pixel density of exported slide images.")
			.addDropdown((dropdown) =>
				dropdown
					.addOption("1", "1× (screen)")
					.addOption("2", "2× (retina)")
					.addOption("3", "3× (print)")
					.setValue(String(this.plugin.settings.exportScale))
					.onChange(async (value) => {
						this.plugin.settings.exportScale = Number(value) || 2;
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl)
			.setName("Export folder")
			.setDesc("Where exported PNGs go. Leave empty to use the vault's attachment folder.")
			.addText((text) =>
				text
					.setPlaceholder("assets/slides")
					.setValue(this.plugin.settings.exportFolder)
					.onChange(async (value) => {
						this.plugin.settings.exportFolder = value.trim();
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl)
			.setName("Include speaker notes")
			.setDesc("Add notes as callouts when extracting a deck to Markdown.")
			.addToggle((toggle) =>
				toggle.setValue(this.plugin.settings.includeNotes).onChange(async (value) => {
					this.plugin.settings.includeNotes = value;
					await this.plugin.saveSettings();
				}),
			);

		new Setting(containerEl)
			.setName("Link back to the deck")
			.setDesc("Add a link to the source .pptx at the top of extracted Markdown.")
			.addToggle((toggle) =>
				toggle.setValue(this.plugin.settings.includeSourceLink).onChange(async (value) => {
					this.plugin.settings.includeSourceLink = value;
					await this.plugin.saveSettings();
				}),
			);

		new Setting(containerEl).setName("Editing").setHeading();
		const info = containerEl.createDiv({ cls: "setting-item-description" });
		info.setText(
			"Double-click slide text to edit it, then save with Cmd/Ctrl+S. The first save of a deck " +
				"leaves a .pptx.bak copy of the original beside it. Everything this plugin does not " +
				"understand — animations, transitions, embedded fonts — is written back unchanged.",
		);
	}
}
