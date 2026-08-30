import { PluginSettingTab, Setting } from "obsidian";
import { t } from "../i18n";
import type PptxStudioPlugin from "../main";

export class PptxSettingsTab extends PluginSettingTab {
	constructor(private readonly plugin: PptxStudioPlugin) {
		super(plugin.app, plugin);
	}

	display(): void {
		const { containerEl } = this;
		containerEl.empty();

		new Setting(containerEl).setName(t("settings.viewer")).setHeading();

		new Setting(containerEl)
			.setName(t("settings.language"))
			.setDesc(t("settings.languageDesc"))
			.addDropdown((dropdown) =>
				dropdown
					.addOption("auto", t("settings.languageAuto"))
					.addOption("en", "English")
					.addOption("ja", "日本語")
					.setValue(this.plugin.settings.language)
					.onChange(async (value) => {
						this.plugin.settings.language =
							value === "en" || value === "ja" ? value : "auto";
						await this.plugin.saveSettings();
						// Rebuild this pane and any open deck so the change is visible now.
						this.display();
						this.plugin.refreshViews();
					}),
			);

		new Setting(containerEl)
			.setName(t("settings.thumbnails"))
			.setDesc(t("settings.thumbnailsDesc"))
			.addToggle((toggle) =>
				toggle.setValue(this.plugin.settings.showThumbnails).onChange(async (value) => {
					this.plugin.settings.showThumbnails = value;
					await this.plugin.saveSettings();
				}),
			);

		new Setting(containerEl)
			.setName(t("settings.notes"))
			.setDesc(t("settings.notesDesc"))
			.addToggle((toggle) =>
				toggle.setValue(this.plugin.settings.showNotes).onChange(async (value) => {
					this.plugin.settings.showNotes = value;
					await this.plugin.saveSettings();
				}),
			);

		new Setting(containerEl)
			.setName(t("settings.zoom"))
			.setDesc(t("settings.zoomDesc"))
			.addDropdown((dropdown) =>
				dropdown
					.addOption("page", t("settings.zoomPage"))
					.addOption("width", t("settings.zoomWidth"))
					.setValue(this.plugin.settings.fitMode)
					.onChange(async (value) => {
						this.plugin.settings.fitMode = value === "width" ? "width" : "page";
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl).setName(t("settings.embeds")).setHeading();

		new Setting(containerEl)
			.setName(t("settings.embedHeight"))
			.setDesc(t("settings.embedHeightDesc"))
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
			.setName(t("settings.embedControls"))
			.setDesc(t("settings.embedControlsDesc"))
			.addToggle((toggle) =>
				toggle.setValue(this.plugin.settings.embedControls).onChange(async (value) => {
					this.plugin.settings.embedControls = value;
					await this.plugin.saveSettings();
				}),
			);

		new Setting(containerEl).setName(t("settings.export")).setHeading();

		new Setting(containerEl)
			.setName(t("settings.pngResolution"))
			.setDesc(t("settings.pngResolutionDesc"))
			.addDropdown((dropdown) =>
				dropdown
					.addOption("1", "1×")
					.addOption("2", "2×")
					.addOption("3", "3×")
					.setValue(String(this.plugin.settings.exportScale))
					.onChange(async (value) => {
						this.plugin.settings.exportScale = Number(value) || 2;
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl)
			.setName(t("settings.exportFolder"))
			.setDesc(t("settings.exportFolderDesc"))
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
			.setName(t("settings.includeNotes"))
			.setDesc(t("settings.includeNotesDesc"))
			.addToggle((toggle) =>
				toggle.setValue(this.plugin.settings.includeNotes).onChange(async (value) => {
					this.plugin.settings.includeNotes = value;
					await this.plugin.saveSettings();
				}),
			);

		new Setting(containerEl)
			.setName(t("settings.linkBack"))
			.setDesc(t("settings.linkBackDesc"))
			.addToggle((toggle) =>
				toggle.setValue(this.plugin.settings.includeSourceLink).onChange(async (value) => {
					this.plugin.settings.includeSourceLink = value;
					await this.plugin.saveSettings();
				}),
			);

		new Setting(containerEl).setName(t("settings.editing")).setHeading();
		containerEl.createDiv({ cls: "setting-item-description", text: t("settings.editingDesc") });
	}
}
