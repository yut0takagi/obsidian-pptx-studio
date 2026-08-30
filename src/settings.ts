import type { LanguageSetting } from "./i18n";

export interface PptxViewerSettings {
	/** Interface language; "auto" follows Obsidian's own setting. */
	language: LanguageSetting;
	/** Show the slide thumbnail rail in the deck view. */
	showThumbnails: boolean;
	/** Show the speaker-notes pane in the deck view. */
	showNotes: boolean;
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

export const DEFAULT_SETTINGS: PptxViewerSettings = {
	language: "auto",
	showThumbnails: true,
	showNotes: false,
	fitMode: "page",
	embedMaxHeight: 420,
	embedControls: true,
	exportScale: 2,
	exportFolder: "",
	includeNotes: true,
	includeSourceLink: true,
};
