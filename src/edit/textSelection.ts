import { runRegistry, textBodyRegistry } from "../render/renderSlide";
import type { Run, TextBody } from "../pptx/types";

export interface SelectedRun {
	run: Run;
	/** Character offsets within the run's text. */
	start: number;
	end: number;
}

export interface TextSelection {
	box: HTMLElement;
	body: TextBody;
	runs: SelectedRun[];
	paragraphs: Element[];
}

/**
 * What the caret has selected inside an open text editor.
 *
 * Formatting from the ribbon means one thing when text is selected — apply to
 * those characters — and another when it is not — apply to the whole shape.
 * Answering that question needs the live DOM selection mapped back onto runs,
 * which is what this does.
 */
export function currentTextSelection(): TextSelection | null {
	const selection = window.getSelection();
	if (!selection || selection.rangeCount === 0 || selection.isCollapsed) return null;
	const range = selection.getRangeAt(0);

	const container =
		range.commonAncestorContainer.instanceOf(HTMLElement)
			? range.commonAncestorContainer
			: range.commonAncestorContainer.parentElement;
	const box = container?.closest<HTMLElement>('[data-editable="1"]') ?? null;
	if (!box) return null;
	const body = textBodyRegistry.get(box);
	if (!body) return null;

	const runs: SelectedRun[] = [];
	for (const span of Array.from(box.querySelectorAll<HTMLElement>("[data-run]"))) {
		if (!range.intersectsNode(span)) continue;
		const run = runRegistry.get(span);
		if (!run?.source) continue;
		const length = (span.textContent ?? "").length;
		const start = span.contains(range.startContainer)
			? offsetWithin(span, range.startContainer, range.startOffset)
			: 0;
		const end = span.contains(range.endContainer)
			? offsetWithin(span, range.endContainer, range.endOffset)
			: length;
		if (end > start) runs.push({ run, start, end });
	}

	const paragraphs: Element[] = [];
	for (const el of Array.from(box.querySelectorAll<HTMLElement>(".pptx-para"))) {
		if (!range.intersectsNode(el)) continue;
		const model = body.paragraphs.find((p) => p.source && paragraphMatches(el, p.source));
		if (model?.source) paragraphs.push(model.source);
	}

	if (runs.length === 0 && paragraphs.length === 0) return null;
	return { box, body, runs, paragraphs };
}

/**
 * The index of the paragraph the caret is in, even with nothing selected.
 *
 * Tab and Shift+Tab change one paragraph's list level, so they need this where a
 * range-based selection would report nothing.
 */
export function caretParagraphIndex(box: HTMLElement): number | null {
	const selection = window.getSelection();
	if (!selection || selection.rangeCount === 0) return null;
	const node = selection.getRangeAt(0).startContainer;
	const el = node.instanceOf(HTMLElement) ? node : node.parentElement;
	const para = el?.closest<HTMLElement>(".pptx-para");
	if (!para || !box.contains(para)) return null;
	const all = Array.from(box.querySelectorAll<HTMLElement>(".pptx-para"));
	const index = all.indexOf(para);
	return index === -1 ? null : index;
}

/** Character offset of a DOM position, counted from the start of `root`. */
function offsetWithin(root: HTMLElement, node: Node, offset: number): number {
	let count = 0;
	let done = false;

	const walk = (current: Node): void => {
		if (done) return;
		if (current === node && current.nodeType !== Node.TEXT_NODE) {
			// An element position counts the text of the children before it.
			for (let i = 0; i < offset && i < current.childNodes.length; i++) {
				count += (current.childNodes[i].textContent ?? "").length;
			}
			done = true;
			return;
		}
		if (current.nodeType === Node.TEXT_NODE) {
			if (current === node) {
				count += offset;
				done = true;
				return;
			}
			count += (current.nodeValue ?? "").length;
			return;
		}
		for (const kid of Array.from(current.childNodes)) {
			walk(kid);
			if (done) return;
		}
	};

	walk(root);
	return count;
}

/**
 * Rendered paragraphs carry no id, so a paragraph element is matched to its
 * model entry through the runs it contains.
 */
function paragraphMatches(el: HTMLElement, source: Element): boolean {
	const span = el.querySelector<HTMLElement>("[data-run]");
	if (!span) return false;
	const run = runRegistry.get(span);
	return run?.source?.parentNode === source;
}
