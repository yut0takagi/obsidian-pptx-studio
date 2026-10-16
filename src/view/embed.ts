import { MarkdownRenderChild, type TFile } from "obsidian";
import { t } from "../i18n";
import { DeckViewer } from "../render/DeckViewer";
import type PptxStudioPlugin from "../main";
import { type SlideRef, parseSlideRef, resolveSlideRef } from "./slideRef";

/**
 * Render `![[deck.pptx]]`, `![[deck.pptx#3]]` and `![[deck.pptx#id=257]]` inside notes.
 *
 * Obsidian only produces an `.internal-embed` span for extensions it knows
 * about, which is why the plugin registers "pptx" before this ever runs.
 */
export function registerPptxEmbeds(plugin: PptxStudioPlugin): void {
	plugin.registerMarkdownPostProcessor((el, ctx) => {
		for (const span of Array.from(el.querySelectorAll<HTMLElement>("span.internal-embed"))) {
			const src = span.getAttribute("src");
			if (!src) continue;
			const [linkpath, subpath] = splitSubpath(src);
			if (!linkpath.toLowerCase().endsWith(".pptx")) continue;

			const file = plugin.app.metadataCache.getFirstLinkpathDest(linkpath, ctx.sourcePath);
			if (!file || file.extension !== "pptx") continue;

			span.empty();
			span.addClass("pptx-embed");
			span.removeClass("internal-embed");
			ctx.addChild(new PptxEmbed(span, plugin, file, parseSlideRef(subpath)));
		}
	});
}

function splitSubpath(src: string): [string, string] {
	const hash = src.indexOf("#");
	return hash === -1 ? [src, ""] : [src.slice(0, hash), src.slice(hash + 1)];
}

class PptxEmbed extends MarkdownRenderChild {
	private viewer: DeckViewer | null = null;

	constructor(
		containerEl: HTMLElement,
		private readonly plugin: PptxStudioPlugin,
		private readonly file: TFile,
		private readonly slideRef: SlideRef | undefined,
	) {
		super(containerEl);
	}

	onload(): void {
		void this.loadDeck();
	}

	private async loadDeck(): Promise<void> {
		const settings = this.plugin.settings;
		this.containerEl.setCssStyles({ height: `${settings.embedMaxHeight}px` });
		const loading = this.containerEl.createDiv({
			cls: "pptx-message",
			text: t("embed.loading", { name: this.file.name }),
		});

		try {
			const loaded = await this.plugin.decks.get(this.file);
			loading.remove();
			if (!this.containerEl.isConnected) return;

			// An id names one slide; showing some other slide in its place would
			// look like it worked, so a missing one says so instead.
			const slideNumber = resolveSlideRef(loaded.deck, this.slideRef);
			if (slideNumber === null) {
				this.containerEl.createDiv({
					cls: "pptx-message pptx-error",
					text: t("embed.slideMissing", { name: this.file.name }),
				});
				this.containerEl.addEventListener("dblclick", this.openInTab);
				return;
			}

			this.viewer = new DeckViewer(this.containerEl, {
				deck: loaded.deck,
				pkg: loaded.pkg,
				compact: true,
				// A one-slide embed needs no chrome; a whole deck needs page controls.
				chrome: settings.embedControls && slideNumber === undefined ? "toolbar" : "none",
				showThumbnails: false,
				showNotes: false,
				fitMode: "page",
				pinnedSlide: slideNumber,
			});

			this.containerEl.addEventListener("dblclick", this.openInTab);
		} catch (error) {
			loading.remove();
			this.containerEl.createDiv({
				cls: "pptx-message pptx-error",
				text: t("embed.failed", { name: this.file.name, message: (error as Error).message }),
			});
		}
	}

	/** Embeds are read-only; double-clicking opens the deck where it can be edited. */
	private openInTab = (event: MouseEvent): void => {
		event.preventDefault();
		void this.plugin.app.workspace.getLeaf("tab").openFile(this.file);
	};

	onunload(): void {
		this.containerEl.removeEventListener("dblclick", this.openInTab);
		this.viewer?.destroy();
		this.viewer = null;
	}
}
