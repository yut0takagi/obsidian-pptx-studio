import { textBodyRegistry } from "../render/renderSlide";
import { caretParagraphIndex } from "./textSelection";
import type { DeckEditor } from "./DeckEditor";
import type { PartsPatch } from "./History";
import { commitTextBody } from "./textEdit";

export interface EditControllerOptions {
	/** Editing is off in embeds and while the file is read-only. */
	isEnabled: () => boolean;
	/** Text edits are recorded here so undo covers them alongside every other edit. */
	editor: DeckEditor;
	/**
	 * Called when a session ends without a change. Editing mutates the live DOM,
	 * so the slide is re-rendered to rebuild the element-to-model registries.
	 */
	onCancelled: () => void;
	/** Tab and Shift+Tab, once the pending text has been committed. */
	onListLevel: (delta: number, target: { shapeId: string; paragraph: number }) => void;
}

/**
 * Click-to-edit for slide text.
 *
 * A text box becomes `contenteditable` in place, so the caret, selection and IME
 * all behave the way the platform expects — which matters a lot for Japanese
 * input, where a custom caret implementation would break composition.
 */
export class EditController {
	private activeBox: HTMLElement | null = null;
	private originalHtml = "";
	private composing = false;
	/** The slide part as it was before this edit session began. */
	private before: PartsPatch | null = null;

	constructor(private readonly options: EditControllerOptions) {}

	get isEditing(): boolean {
		return this.activeBox !== null;
	}

	/** Wire up a freshly rendered slide element. */
	attach(slideEl: HTMLElement): void {
		slideEl.addEventListener("dblclick", this.onDoubleClick);
	}

	detach(slideEl: HTMLElement): void {
		slideEl.removeEventListener("dblclick", this.onDoubleClick);
	}

	private onDoubleClick = (event: MouseEvent): void => {
		if (!this.options.isEnabled()) return;
		const target = event.target;
		if (!(target instanceof HTMLElement)) return;
		const box = target.closest<HTMLElement>('[data-editable="1"]');
		if (!box) return;
		event.preventDefault();
		event.stopPropagation();
		if (box === this.activeBox) return;
		this.commit();
		this.begin(box, event);
	};

	/**
	 * Start editing without a click, for Enter and F2. The caret goes to the end
	 * of the text, which is where someone pressing Enter to "start typing" means.
	 */
	beginAtEnd(box: HTMLElement): void {
		if (!this.options.isEnabled() || box === this.activeBox) return;
		this.commit();
		this.begin(box, null);
		const selection = window.getSelection();
		if (!selection) return;
		const range = document.createRange();
		range.selectNodeContents(box);
		range.collapse(false);
		selection.removeAllRanges();
		selection.addRange(range);
	}

	private begin(box: HTMLElement, event: MouseEvent | null): void {
		this.activeBox = box;
		this.originalHtml = box.innerHTML;
		const part = textBodyRegistry.get(box)?.sourcePart;
		this.before = part ? this.options.editor.capture([part]) : null;
		box.contentEditable = "true";
		box.spellcheck = false;
		box.addClass("is-editing");
		// Text can overflow its shape while typing; let it be seen rather than clipped.
		box.style.overflow = "visible";

		box.addEventListener("keydown", this.onKeyDown);
		box.addEventListener("focusout", this.onFocusOut);
		box.addEventListener("compositionstart", this.onCompositionStart);
		box.addEventListener("compositionend", this.onCompositionEnd);

		box.focus({ preventScroll: true });
		if (event) this.placeCaret(event);
	}

	/** Put the caret where the user double-clicked rather than at the start. */
	private placeCaret(event: MouseEvent): void {
		const doc = document as Document & {
			caretRangeFromPoint?: (x: number, y: number) => Range | null;
		};
		const range = doc.caretRangeFromPoint?.(event.clientX, event.clientY);
		if (!range) return;
		const selection = window.getSelection();
		if (!selection) return;
		selection.removeAllRanges();
		selection.addRange(range);
	}

	private onCompositionStart = (): void => {
		this.composing = true;
	};

	private onCompositionEnd = (): void => {
		this.composing = false;
	};

	private onKeyDown = (event: KeyboardEvent): void => {
		// Never let slide navigation or zoom shortcuts fire while typing.
		event.stopPropagation();
		if (this.composing) return;
		if (event.key === "Escape") {
			event.preventDefault();
			this.cancel();
			return;
		}
		if (event.key === "Enter" && (event.metaKey || event.ctrlKey)) {
			event.preventDefault();
			this.commit();
			return;
		}
		if (event.key === "Tab") {
			event.preventDefault();
			const box = this.activeBox;
			if (!box) return;
			const paragraph = caretParagraphIndex(box);
			const shapeId = ownerShapeId(box);
			this.commit();
			if (paragraph !== null && shapeId !== null) {
				this.options.onListLevel(event.shiftKey ? -1 : 1, { shapeId, paragraph });
			}
		}
	};

	private onFocusOut = (event: FocusEvent): void => {
		const next = event.relatedTarget;
		if (next instanceof Node && this.activeBox?.contains(next)) return;
		this.commit();
	};

	/** Write the active editor back to XML. Returns true if anything changed. */
	commit(): boolean {
		const box = this.activeBox;
		if (!box) return false;
		const before = this.before;
		this.teardown(box);
		const result = commitTextBody(box);
		if (result.changed && result.part && before) {
			this.options.editor.recordApplied("Edit text", before, true);
		} else {
			this.options.onCancelled();
		}
		return result.changed;
	}

	/** Abandon the active edit, restoring what was on screen before it started. */
	cancel(): void {
		const box = this.activeBox;
		if (!box) return;
		box.innerHTML = this.originalHtml;
		this.teardown(box);
		this.options.onCancelled();
	}

	private teardown(box: HTMLElement): void {
		box.contentEditable = "false";
		box.removeClass("is-editing");
		box.style.overflow = "";
		box.removeEventListener("keydown", this.onKeyDown);
		box.removeEventListener("focusout", this.onFocusOut);
		box.removeEventListener("compositionstart", this.onCompositionStart);
		box.removeEventListener("compositionend", this.onCompositionEnd);
		this.activeBox = null;
		this.originalHtml = "";
		this.composing = false;
		this.before = null;
	}
}

/** The id of the shape a text box belongs to, read from its p:cNvPr. */
function ownerShapeId(box: HTMLElement): string | null {
	const sp = textBodyRegistry.get(box)?.source?.parentNode;
	if (!(sp instanceof Element)) return null;
	for (const node of Array.from(sp.getElementsByTagName("*"))) {
		if (node.localName === "cNvPr") return node.getAttribute("id");
	}
	return null;
}
