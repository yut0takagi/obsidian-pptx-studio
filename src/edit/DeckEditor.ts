import type { PptxPackage } from "../pptx/package";
import { History, type PartsPatch } from "./History";

export interface DeckEditorOptions {
	/**
	 * Called after a successful edit. `rebuild` is true when the model has to be
	 * re-derived — either the shape tree changed, or XML nodes were replaced and
	 * the model's element references are stale. `parts` names what was touched,
	 * so the listener can re-derive one slide rather than the deck.
	 */
	onChanged: (rebuild: boolean, parts: string[]) => void;
}

/**
 * The single point every mutation goes through.
 *
 * Having one place that snapshots, mutates, records and notifies means undo can
 * never be forgotten for a new command, and the difference between a cheap edit
 * (drag a shape: mutate in place, keep the selection) and an expensive one
 * (delete a shape: the tree changed, re-render) is one flag rather than a
 * convention each command has to remember.
 */
export class DeckEditor {
	readonly history = new History();

	constructor(
		private readonly pkg: PptxPackage,
		private readonly options: DeckEditorOptions,
	) {}

	/**
	 * Apply a mutation to the named parts, recording it for undo.
	 *
	 * `mutate` returns false to abandon the edit — nothing is recorded and no
	 * change is reported, so a command that turns out to be a no-op costs nothing.
	 */
	transact(
		label: string,
		parts: string[],
		mutate: () => boolean,
		options: { rebuild?: boolean } = {},
	): boolean {
		const unique = [...new Set(parts)];
		const before: PartsPatch = {};
		for (const path of unique) before[path] = this.pkg.serializePart(path);

		let ok = false;
		try {
			ok = mutate();
		} catch (error) {
			// A half-applied edit is worse than none: put the parts back.
			for (const [path, bytes] of Object.entries(before)) this.pkg.replacePart(path, bytes);
			throw error;
		}
		if (!ok) return false;

		const after: PartsPatch = {};
		for (const path of unique) {
			after[path] = this.pkg.serializePart(path);
			if (after[path] !== null) this.pkg.markDirty(path);
		}

		this.history.record({ label, before, after });
		this.options.onChanged(options.rebuild ?? true, unique);
		return true;
	}

	/**
	 * Record an edit that has already been applied, given the bytes captured
	 * before it. Used by drag interactions, which must update the live DOM as the
	 * pointer moves and only become an undoable step when it is released.
	 */
	recordApplied(label: string, before: PartsPatch, rebuild = false): void {
		const after: PartsPatch = {};
		for (const path of Object.keys(before)) {
			after[path] = this.pkg.serializePart(path);
			if (after[path] !== null) this.pkg.markDirty(path);
		}
		this.history.record({ label, before, after });
		this.options.onChanged(rebuild, Object.keys(before));
	}

	/** Capture the current bytes of parts, for use with recordApplied. */
	capture(parts: string[]): PartsPatch {
		const patch: PartsPatch = {};
		for (const path of [...new Set(parts)]) patch[path] = this.pkg.serializePart(path);
		return patch;
	}

	undo(): boolean {
		const entry = this.history.undo(this.pkg);
		if (!entry) return false;
		this.options.onChanged(true, Object.keys(entry.before));
		return true;
	}

	redo(): boolean {
		const entry = this.history.redo(this.pkg);
		if (!entry) return false;
		this.options.onChanged(true, Object.keys(entry.after));
		return true;
	}

	get canUndo(): boolean {
		return this.history.canUndo;
	}

	get canRedo(): boolean {
		return this.history.canRedo;
	}
}
