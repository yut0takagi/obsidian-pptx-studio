import { MarkdownRenderChild, type TFile } from "obsidian";
import { DeckViewer } from "../render/DeckViewer";
import type PptxViewerPlugin from "../main";

/**
 * Render `![[deck.pptx]]` and `![[deck.pptx#3]]` inside notes.
 *
 * Obsidian only produces an `.internal-embed` span for extensions it knows
 * about, which is why the plugin registers "pptx" before this ever runs.
 */
export function registerPptxEmbeds(plugin: PptxViewerPlugin): void {
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
			ctx.addChild(new PptxEmbed(span, plugin, file, parseSlideNumber(subpath)));
		}
	});
}

function splitSubpath(src: string): [string, string] {
	const hash = src.indexOf("#");
	return hash === -1 ? [src, ""] : [src.slice(0, hash), src.slice(hash + 1)];
}

/** `#3` and `#slide=3` both pin the embed to slide 3. */
function parseSlideNumber(subpath: string): number | undefined {
	const match = /^(?:slide=?)?(\d+)$/i.exec(subpath.trim());
	if (!match) return undefined;
	const n = Number(match[1]);
	return Number.isFinite(n) && n > 0 ? n : undefined;
}

class PptxEmbed extends MarkdownRenderChild {
	private viewer: DeckViewer | null = null;

	constructor(
		containerEl: HTMLElement,
		private readonly plugin: PptxViewerPlugin,
		private readonly file: TFile,
		private readonly slideNumber: number | undefined,
	) {
		super(containerEl);
	}

	async onload(): Promise<void> {
		const settings = this.plugin.settings;
		this.containerEl.style.height = `${settings.embedMaxHeight}px`;
		const loading = this.containerEl.createDiv({
			cls: "pptx-message",
			text: `Loading ${this.file.name}…`,
		});

		try {
			const loaded = await this.plugin.decks.get(this.file);
			loading.remove();
			if (!this.containerEl.isConnected) return;

			this.viewer = new DeckViewer(this.containerEl, {
				deck: loaded.deck,
				pkg: loaded.pkg,
				compact: true,
				// A one-slide embed needs no chrome; a whole deck needs page controls.
				chrome:
					settings.embedControls && this.slideNumber === undefined ? "toolbar" : "none",
				showThumbnails: false,
				showNotes: false,
				fitMode: "page",
				pinnedSlide: this.slideNumber,
			});

			this.containerEl.addEventListener("dblclick", this.openInTab);
		} catch (error) {
			loading.remove();
			this.containerEl.createDiv({
				cls: "pptx-message pptx-error",
				text: `Could not render ${this.file.name}: ${(error as Error).message}`,
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
