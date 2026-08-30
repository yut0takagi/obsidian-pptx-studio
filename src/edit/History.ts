/**
 * Undo/redo for deck edits.
 *
 * Every edit is a mutation of one shape's XML subtree, so a snapshot of that
 * subtree is a complete, self-contained record of it. That keeps one mechanism
 * covering both text and geometry, and means undo can never leave the XML in a
 * state the parser has not seen.
 */
export interface Snapshot {
	target: Element;
	attributes: [string, string][];
	children: Node[];
}

export interface HistoryEntry {
	label: string;
	/** Package part the snapshot's element lives in. */
	part: string;
	before: Snapshot;
	after: Snapshot;
}

export function capture(target: Element): Snapshot {
	return {
		target,
		attributes: Array.from(target.attributes).map((a) => [a.name, a.value] as [string, string]),
		children: Array.from(target.childNodes).map((n) => n.cloneNode(true)),
	};
}

function restore(snapshot: Snapshot): void {
	const { target } = snapshot;
	while (target.firstChild) target.removeChild(target.firstChild);
	for (const name of Array.from(target.attributes).map((a) => a.name)) {
		target.removeAttribute(name);
	}
	for (const [name, value] of snapshot.attributes) target.setAttribute(name, value);
	for (const node of snapshot.children) target.appendChild(node.cloneNode(true));
}

export class ElementHistory {
	private readonly past: HistoryEntry[] = [];
	private readonly future: HistoryEntry[] = [];

	constructor(private readonly limit = 100) {}

	/** Record an edit, given the snapshot taken before it was applied. */
	record(label: string, part: string, before: Snapshot): void {
		this.past.push({ label, part, before, after: capture(before.target) });
		this.future.length = 0;
		if (this.past.length > this.limit) this.past.shift();
	}

	get canUndo(): boolean {
		return this.past.length > 0;
	}

	get canRedo(): boolean {
		return this.future.length > 0;
	}

	/** Undo the last edit. Returns the part that changed, or null if there was none. */
	undo(): HistoryEntry | null {
		const entry = this.past.pop();
		if (!entry) return null;
		restore(entry.before);
		this.future.push(entry);
		return entry;
	}

	redo(): HistoryEntry | null {
		const entry = this.future.pop();
		if (!entry) return null;
		restore(entry.after);
		this.past.push(entry);
		return entry;
	}

	clear(): void {
		this.past.length = 0;
		this.future.length = 0;
	}
}
