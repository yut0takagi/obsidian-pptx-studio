import type { LanguageSetting } from "./i18n";

export interface PptxStudioSettings {
	/** Interface language; "auto" follows Obsidian's own setting. */
	language: LanguageSetting;
	/** Ribbon state, remembered between sessions. */
	ribbonCollapsed: boolean;
	ribbonTab: string;
	/** Show the slide thumbnail rail in the deck view. */
	showThumbnails: boolean;
	/** Show the speaker-notes pane in the deck view. */
	showNotes: boolean;
	/** Width, in pixels, of the side pane holding the shape list and inspector. */
	sidePaneWidth: number;
	/** Height, in pixels, of the shape list above the inspector. */
	shapeListHeight: number;
	/** Whether either side-pane section is folded away. */
	shapeListCollapsed: boolean;
	inspectorCollapsed: boolean;
	/** How a slide is sized when the view opens or the pane resizes. */
	fitMode: "page" | "width";
	/** Maximum height, in pixels, of a slide embedded in a note. */
	embedMaxHeight: number;
	/** Show navigation controls on embeds that span the whole deck. */
	embedControls: boolean;
	/** Pixel density multiplier for exported PNGs. */
	exportScale: number;
	/** Folder for exported PNGs; empty means the vault's attachment folder. */
	exportFolder: string;
	/** Include speaker notes when extracting a deck to Markdown. */
	includeNotes: boolean;
	/** Include a link back to the source .pptx in extracted Markdown. */
	includeSourceLink: boolean;
}

export const DEFAULT_SETTINGS: PptxStudioSettings = {
	language: "auto",
	ribbonCollapsed: false,
	ribbonTab: "home",
	showThumbnails: true,
	showNotes: false,
	sidePaneWidth: 260,
	shapeListHeight: 180,
	shapeListCollapsed: false,
	inspectorCollapsed: false,
	fitMode: "page",
	embedMaxHeight: 420,
	embedControls: true,
	exportScale: 2,
	exportFolder: "",
	includeNotes: true,
	includeSourceLink: true,
};
