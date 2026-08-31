import { textBodyRegistry } from "../render/renderSlide";
import { CaretPainter } from "./CaretPainter";
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
	/** Editing started or ended, so the selection can be drawn differently. */
	onModeChange?: (editing: boolean) => void;
	/**
	 * The user asked to go back to holding the shape — Escape, or Cmd+Enter. The
	 * canvas takes the keyboard again here; leaving focus on the box that is
	 * about to be re-rendered would strand the arrow keys.
	 */
	onReturnToShape?: () => void;
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
	private originalNodes: Node[] = [];
	private composing = false;
	private readonly caret = new CaretPainter();
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
		if (!isHtmlElement(target)) return;
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
		this.selectAllText(box, true);
	}

	/**
	 * Start editing by typing over the shape, which replaces what was there.
	 *
	 * Selecting a shape and typing means "this shape says that now" — the same
	 * gesture as in PowerPoint, and the reason typing is not simply ignored while
	 * a shape rather than its text is selected.
	 */
	beginReplacing(box: HTMLElement, initial: string): void {
		if (!this.options.isEnabled() || box === this.activeBox) return;
		this.commit();
		this.begin(box, null);
		this.selectAllText(box, false);
		document.execCommand("insertText", false, initial);
	}

	private selectAllText(box: HTMLElement, collapseToEnd: boolean): void {
		const selection = window.getSelection();
		if (!selection) return;
		const range = document.createRange();
		range.selectNodeContents(box);
		if (collapseToEnd) range.collapse(false);
		selection.removeAllRanges();
		selection.addRange(range);
	}

	private begin(box: HTMLElement, event: MouseEvent | null): void {
		this.activeBox = box;
		this.originalNodes = Array.from(box.childNodes).map((node) => node.cloneNode(true));
		const part = textBodyRegistry.get(box)?.sourcePart;
		this.before = part ? this.options.editor.capture([part]) : null;
		box.contentEditable = "true";
		box.spellcheck = false;
		box.addClass("is-editing");
		// Text can overflow its shape while typing; let it be seen rather than clipped.
		box.setCssStyles({ overflow: "visible" });

		box.addEventListener("keydown", this.onKeyDown);
		box.addEventListener("focusout", this.onFocusOut);
		this.options.onModeChange?.(true);
		box.addEventListener("compositionstart", this.onCompositionStart);
		box.addEventListener("compositionend", this.onCompositionEnd);

		box.focus({ preventScroll: true });
		if (event) this.placeCaret(event);
		this.caret.start(box);
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
			// PowerPoint's Escape leaves text editing and keeps the shape selected;
			// it does not throw the typing away, which is what a revert here would do.
			event.preventDefault();
			this.commit();
			this.options.onReturnToShape?.();
			return;
		}
		if (event.key === "Enter" && (event.metaKey || event.ctrlKey)) {
			event.preventDefault();
			this.commit();
			this.options.onReturnToShape?.();
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
		box.empty();
		box.append(...this.originalNodes.map((node) => node.cloneNode(true)));
		this.teardown(box);
		this.options.onCancelled();
	}

	private teardown(box: HTMLElement): void {
		this.caret.stop();
		box.contentEditable = "false";
		box.removeClass("is-editing");
		box.setCssStyles({ overflow: "" });
		box.removeEventListener("keydown", this.onKeyDown);
		box.removeEventListener("focusout", this.onFocusOut);
		box.removeEventListener("compositionstart", this.onCompositionStart);
		box.removeEventListener("compositionend", this.onCompositionEnd);
		this.options.onModeChange?.(false);
		this.activeBox = null;
		this.originalNodes = [];
		this.composing = false;
		this.before = null;
	}
}

/** The id of the shape a text box belongs to, read from its p:cNvPr. */
function ownerShapeId(box: HTMLElement): string | null {
	const sp = textBodyRegistry.get(box)?.source?.parentNode;
	if (!sp?.instanceOf(Element)) return null;
	for (const node of Array.from(sp.getElementsByTagName("*"))) {
		if (node.localName === "cNvPr") return node.getAttribute("id");
	}
	return null;
}

function isHtmlElement(value: EventTarget | null): value is HTMLElement {
	return (value as Node | null)?.instanceOf(HTMLElement) === true;
}
