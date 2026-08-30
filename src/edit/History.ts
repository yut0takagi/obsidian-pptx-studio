import type { PptxPackage } from "../pptx/package";

/**
 * A set of package parts and the bytes they should hold. `null` means the part
 * is absent, which is how part creation and deletion round trip.
 */
export type PartsPatch = Record<string, Uint8Array | null>;

export interface HistoryEntry {
	label: string;
	before: PartsPatch;
	after: PartsPatch;
}

/**
 * Undo/redo at the level of whole package parts.
 *
 * An earlier version snapshotted the edited element's XML subtree, which is
 * cheaper but cannot express an edit that adds or removes a part — inserting a
 * picture, adding a slide. Recording parts instead covers every edit the editor
 * can make with one mechanism, and guarantees undo lands on a state the parser
 * has already accepted.
 */
export class History {
	private readonly past: HistoryEntry[] = [];
	private readonly future: HistoryEntry[] = [];

	constructor(private readonly limit = 60) {}

	record(entry: HistoryEntry): void {
		this.past.push(entry);
		this.future.length = 0;
		if (this.past.length > this.limit) this.past.shift();
	}

	get canUndo(): boolean {
		return this.past.length > 0;
	}

	get canRedo(): boolean {
		return this.future.length > 0;
	}

	get undoLabel(): string | null {
		return this.past.at(-1)?.label ?? null;
	}

	get redoLabel(): string | null {
		return this.future.at(-1)?.label ?? null;
	}

	undo(pkg: PptxPackage): HistoryEntry | null {
		const entry = this.past.pop();
		if (!entry) return null;
		applyPatch(pkg, entry.before);
		this.future.push(entry);
		return entry;
	}

	redo(pkg: PptxPackage): HistoryEntry | null {
		const entry = this.future.pop();
		if (!entry) return null;
		applyPatch(pkg, entry.after);
		this.past.push(entry);
		return entry;
	}

	clear(): void {
		this.past.length = 0;
		this.future.length = 0;
	}
}

function applyPatch(pkg: PptxPackage, patch: PartsPatch): void {
	for (const [path, bytes] of Object.entries(patch)) {
		pkg.replacePart(path, bytes);
		if (bytes !== null) pkg.markDirty(path);
	}
}
