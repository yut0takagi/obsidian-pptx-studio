import { PluginSettingTab, Setting, type SettingDefinitionItem } from "obsidian";
import { t } from "../i18n";
import type PptxStudioPlugin from "../main";
import type { PptxStudioSettings } from "../settings";

export class PptxSettingsTab extends PluginSettingTab {
	constructor(private readonly plugin: PptxStudioPlugin) {
		super(plugin.app, plugin);
	}

	getSettingDefinitions(): SettingDefinitionItem<keyof PptxStudioSettings>[] {
		return [
			{
				type: "group",
				heading: t("settings.viewer"),
				items: [
					{
						name: t("settings.language"),
						desc: t("settings.languageDesc"),
						control: {
							type: "dropdown",
							key: "language",
							options: {
								auto: t("settings.languageAuto"),
								en: "English",
								ja: "日本語",
							},
						},
					},
					{
						name: t("settings.thumbnails"),
						desc: t("settings.thumbnailsDesc"),
						control: { type: "toggle", key: "showThumbnails" },
					},
					{
						name: t("settings.notes"),
						desc: t("settings.notesDesc"),
						control: { type: "toggle", key: "showNotes" },
					},
					{
						name: t("settings.zoom"),
						desc: t("settings.zoomDesc"),
						control: {
							type: "dropdown",
							key: "fitMode",
							options: {
								page: t("settings.zoomPage"),
								width: t("settings.zoomWidth"),
							},
						},
					},
				],
			},
			{
				type: "group",
				heading: t("settings.embeds"),
				items: [
					{
						name: t("settings.embedHeight"),
						desc: t("settings.embedHeightDesc"),
						control: {
							type: "slider",
							key: "embedMaxHeight",
							min: 200,
							max: 900,
							step: 20,
						},
					},
					{
						name: t("settings.embedControls"),
						desc: t("settings.embedControlsDesc"),
						control: { type: "toggle", key: "embedControls" },
					},
				],
			},
			{
				type: "group",
				heading: t("settings.export"),
				items: [
					{
						name: t("settings.pngResolution"),
						desc: t("settings.pngResolutionDesc"),
						control: {
							type: "dropdown",
							key: "exportScale",
							options: {
								"1": "1×",
								"2": "2×",
								"3": "3×",
							},
						},
					},
					{
						name: t("settings.exportFolder"),
						desc: t("settings.exportFolderDesc"),
						control: {
							type: "text",
							key: "exportFolder",
							placeholder: "assets/slides",
						},
					},
					{
						name: t("settings.includeNotes"),
						desc: t("settings.includeNotesDesc"),
						control: { type: "toggle", key: "includeNotes" },
					},
					{
						name: t("settings.linkBack"),
						desc: t("settings.linkBackDesc"),
						control: { type: "toggle", key: "includeSourceLink" },
					},
				],
			},
			{
				type: "group",
				heading: t("settings.editing"),
				items: [{ name: t("settings.editing"), desc: t("settings.editingDesc") }],
			},
		];
	}

	getControlValue(key: keyof PptxStudioSettings): unknown {
		if (key === "exportScale") return String(this.plugin.settings.exportScale);
		return this.plugin.settings[key];
	}

	async setControlValue(key: keyof PptxStudioSettings, value: unknown): Promise<void> {
		switch (key) {
			case "language":
				this.plugin.settings.language =
					value === "en" || value === "ja" ? value : "auto";
				await this.plugin.saveSettings();
				this.update();
				this.plugin.refreshViews();
				return;
			case "showThumbnails":
			case "showNotes":
			case "embedControls":
			case "includeNotes":
			case "includeSourceLink":
				this.plugin.settings[key] = value === true;
				break;
			case "fitMode":
				this.plugin.settings.fitMode = value === "width" ? "width" : "page";
				break;
			case "embedMaxHeight":
				this.plugin.settings.embedMaxHeight = typeof value === "number" ? value : 420;
				break;
			case "exportScale":
				this.plugin.settings.exportScale = Number(value) || 2;
				break;
			case "exportFolder":
				this.plugin.settings.exportFolder = String(value ?? "").trim();
				break;
			default:
				return;
		}
		await this.plugin.saveSettings();
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
