import type { App, TFile } from "obsidian";
import type { PptxPackage } from "./pptx/package";
import { parseDeck } from "./pptx/parse";
import type { Deck } from "./pptx/types";

export interface LoadedDeck {
	deck: Deck;
	pkg: PptxPackage;
}

interface Entry extends LoadedDeck {
	path: string;
	mtime: number;
	size: number;
	lastUsed: number;
}

/**
 * Parsed decks, shared by the file view and by every embed in every note.
 *
 * A deck owns object URLs for its images, so its lifetime has to outlive any
 * view showing it. Keeping ownership here — rather than in a view — means an
 * embed and a full view of the same file share one parse, and closing either
 * one does not pull images out from under the other.
 */
export class DeckCache {
	private readonly entries = new Map<string, Entry>();
	private readonly inFlight = new Map<string, Promise<LoadedDeck>>();

	constructor(
		private readonly app: App,
		private readonly maxEntries = 4,
	) {}

	async get(file: TFile): Promise<LoadedDeck> {
		const existing = this.entries.get(file.path);
		if (existing && existing.mtime === file.stat.mtime && existing.size === file.stat.size) {
			existing.lastUsed = Date.now();
			return existing;
		}
		// An edited deck is newer than what is on disk; keep it rather than reloading.
		if (existing?.pkg.isDirty) {
			existing.lastUsed = Date.now();
			return existing;
		}
		if (existing) this.evict(file.path);

		const pending = this.inFlight.get(file.path);
		if (pending) return pending;

		const task = (async () => {
			const data = await this.app.vault.readBinary(file);
			const loaded = parseDeck(data, file.basename);
			// The cache entry *is* the object handed to callers, so a view that
			// swaps in a rebuilt deck after an edit updates the cache too.
			const entry: Entry = {
				deck: loaded.deck,
				pkg: loaded.pkg,
				path: file.path,
				mtime: file.stat.mtime,
				size: file.stat.size,
				lastUsed: Date.now(),
			};
			this.entries.set(file.path, entry);
			this.trim();
			return entry;
		})();

		this.inFlight.set(file.path, task);
		try {
			return await task;
		} finally {
			this.inFlight.delete(file.path);
		}
	}

	/** Drop a deck, releasing its object URLs. Unsaved edits are never discarded. */
	evict(path: string, force = false): void {
		const entry = this.entries.get(path);
		if (!entry) return;
		if (entry.pkg.isDirty && !force) return;
		this.entries.delete(path);
		entry.pkg.dispose();
	}

	/** Record the on-disk stats a deck now matches, e.g. straight after a save. */
	touch(path: string, mtime: number, size: number): void {
		const entry = this.entries.get(path);
		if (!entry) return;
		entry.mtime = mtime;
		entry.size = size;
		entry.lastUsed = Date.now();
	}

	rename(oldPath: string, newPath: string): void {
		const entry = this.entries.get(oldPath);
		if (!entry) return;
		this.entries.delete(oldPath);
		entry.path = newPath;
		this.entries.set(newPath, entry);
	}

	clear(): void {
		for (const path of Array.from(this.entries.keys())) this.evict(path, true);
	}

	/** True when a deck is open with edits that have not been written to disk. */
	hasUnsavedChanges(path: string): boolean {
		return this.entries.get(path)?.pkg.isDirty ?? false;
	}

	private trim(): void {
		while (this.entries.size > this.maxEntries) {
			let oldest: Entry | null = null;
			for (const entry of this.entries.values()) {
				// A deck with unsaved edits stays resident however old it is.
				if (entry.pkg.isDirty) continue;
				if (!oldest || entry.lastUsed < oldest.lastUsed) oldest = entry;
			}
			if (!oldest) return;
			this.evict(oldest.path);
		}
	}
}
