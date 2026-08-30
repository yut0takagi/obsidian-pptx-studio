/**
 * What is selected on the current slide.
 *
 * Selection is stored as shape ids rather than objects or elements, so it
 * survives the model being rebuilt after an edit. Deleting one of three
 * selected shapes, or undoing, leaves the rest still selected — which is what
 * makes a sequence of edits feel continuous rather than resetting each time.
 */
export class Selection {
	private slide = -1;
	private members = new Set<string>();
	private readonly listeners = new Set<() => void>();

	get slideIndex(): number {
		return this.slide;
	}

	get ids(): ReadonlySet<string> {
		return this.members;
	}

	get size(): number {
		return this.members.size;
	}

	get isEmpty(): boolean {
		return this.members.size === 0;
	}

	has(id: string): boolean {
		return this.members.has(id);
	}

	set(slideIndex: number, ids: Iterable<string>): void {
		const next = new Set(ids);
		if (slideIndex === this.slide && sameSet(next, this.members)) return;
		this.slide = slideIndex;
		this.members = next;
		this.notify();
	}

	add(slideIndex: number, id: string): void {
		if (slideIndex !== this.slide) {
			this.set(slideIndex, [id]);
			return;
		}
		if (this.members.has(id)) return;
		this.members.add(id);
		this.notify();
	}

	toggle(slideIndex: number, id: string): void {
		if (slideIndex !== this.slide) {
			this.set(slideIndex, [id]);
			return;
		}
		if (this.members.has(id)) this.members.delete(id);
		else this.members.add(id);
		this.notify();
	}

	remove(id: string): void {
		if (!this.members.delete(id)) return;
		this.notify();
	}

	clear(): void {
		if (this.members.size === 0) return;
		this.members.clear();
		this.notify();
	}

	/** Drop ids that no longer exist, e.g. after a delete or an undo. */
	retain(existing: Set<string>): void {
		let changed = false;
		for (const id of [...this.members]) {
			if (!existing.has(id)) {
				this.members.delete(id);
				changed = true;
			}
		}
		if (changed) this.notify();
	}

	onChange(listener: () => void): () => void {
		this.listeners.add(listener);
		return () => this.listeners.delete(listener);
	}

	private notify(): void {
		for (const listener of this.listeners) listener();
	}
}

function sameSet(a: Set<string>, b: Set<string>): boolean {
	if (a.size !== b.size) return false;
	for (const value of a) if (!b.has(value)) return false;
	return true;
}
