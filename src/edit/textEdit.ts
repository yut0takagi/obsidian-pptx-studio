import { paragraphRegistry, runRegistry, textBodyRegistry } from "../render/renderSlide";
import type { Paragraph, Run, TextBody } from "../pptx/types";
import { child, children } from "../pptx/xml";

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

/** One paragraph as it exists in the edited DOM. */
interface EditedParagraph {
	/** The model paragraph this block came from, or null if the user created it. */
	model: Paragraph | null;
	text: string;
	runs: { run: Run; text: string }[];
}

export interface CommitResult {
	changed: boolean;
	/** Package part that needs re-serialising, when something changed. */
	part: string | null;
}

/**
 * Write an edited text box back into the deck's XML.
 *
 * Runs are matched one-to-one where they survived the edit, so typing inside a
 * bold word keeps it bold. Only when the user restructures a paragraph — merging
 * runs, deleting across them — does the paragraph collapse onto the formatting
 * of its first run, which is the behaviour that surprises people least.
 */
export function commitTextBody(boxEl: HTMLElement): CommitResult {
	const body = textBodyRegistry.get(boxEl);
	if (!body?.source) return { changed: false, part: null };

	const edited = collectParagraphs(boxEl);
	if (!hasChanges(body, edited)) return { changed: false, part: null };

	const txBody = body.source;
	const originals = body.paragraphs
		.map((p) => p.source)
		.filter((el): el is Element => el !== null);
	const kept = new Set<Element>();
	let previous: Element | null = null;

	for (const entry of edited) {
		const source = entry.model?.source ?? null;
		if (source && source.parentNode === txBody) {
			applyParagraph(source, entry);
			kept.add(source);
			previous = source;
			continue;
		}
		// A paragraph the user created with Enter: clone the one above it so the
		// new line inherits its list level, alignment and character formatting.
		const template = previous ?? originals[0] ?? null;
		if (!template) continue;
		const clone = template.cloneNode(true) as Element;
		collapseToSingleRun(clone, entry.text);
		txBody.insertBefore(clone, previous ? previous.nextSibling : firstParagraph(txBody));
		previous = clone;
	}

	for (const el of originals) {
		if (!kept.has(el) && el.parentNode === txBody) txBody.removeChild(el);
	}

	return { changed: true, part: body.sourcePart };
}

function firstParagraph(txBody: Element): Element | null {
	return child(txBody, "p");
}

function hasChanges(body: TextBody, edited: EditedParagraph[]): boolean {
	if (edited.length !== body.paragraphs.length) return true;
	for (let i = 0; i < edited.length; i++) {
		const original = body.paragraphs[i].runs.map((r) => r.text).join("");
		if (edited[i].text !== original) return true;
	}
	return false;
}

function applyParagraph(paraEl: Element, entry: EditedParagraph): void {
	const model = entry.model;
	const modelRuns = model ? model.runs.filter((r) => r.source) : [];
	const editedRuns = entry.runs.filter((r) => r.run.source);

	const sameShape =
		modelRuns.length === editedRuns.length &&
		editedRuns.every((r, i) => r.run.source === modelRuns[i].source);

	if (sameShape && modelRuns.length > 0) {
		for (const { run, text } of editedRuns) {
			// a:br carries no text of its own; leaving it alone preserves the break.
			if (!run.source || run.source.localName === "br") continue;
			setRunText(run.source, text);
		}
		return;
	}

	collapseToSingleRun(paraEl, entry.text);
}

function setRunText(runEl: Element, text: string): void {
	const t = child(runEl, "t");
	if (t) {
		t.textContent = text;
		// Leading and trailing spaces only survive with xml:space="preserve".
		if (/^\s|\s$/.test(text)) t.setAttribute("xml:space", "preserve");
		else t.removeAttribute("xml:space");
		return;
	}
	const created = runEl.ownerDocument.createElementNS(A_NS, "a:t");
	created.textContent = text;
	runEl.appendChild(created);
}

/**
 * Replace every run in a paragraph with a single run carrying `text`, keeping
 * the first run's character formatting (or a:endParaRPr's when there is none).
 */
function collapseToSingleRun(paraEl: Element, text: string): void {
	const doc = paraEl.ownerDocument;
	const runLike = ["r", "fld", "br"].flatMap((name) => children(paraEl, name));
	const firstRun = runLike.find((el) => el.localName === "r" || el.localName === "fld") ?? null;

	if (text === "") {
		for (const el of runLike) paraEl.removeChild(el);
		return;
	}

	let target: Element;
	if (firstRun && firstRun.localName === "r") {
		target = firstRun;
	} else {
		target = doc.createElementNS(A_NS, "a:r");
		const rPr = firstRun ? child(firstRun, "rPr") : child(paraEl, "endParaRPr");
		if (rPr) {
			const copy = doc.createElementNS(A_NS, "a:rPr");
			for (const a of Array.from(rPr.attributes)) copy.setAttribute(a.name, a.value);
			for (const kid of Array.from(rPr.childNodes)) copy.appendChild(kid.cloneNode(true));
			target.appendChild(copy);
		}
		const pPr = child(paraEl, "pPr");
		paraEl.insertBefore(target, pPr ? pPr.nextSibling : paraEl.firstChild);
	}

	setRunText(target, text);
	for (const el of runLike) {
		if (el !== target) paraEl.removeChild(el);
	}
}

// ------------------------------------------------------- reading the DOM

/** Walk the edited DOM into a flat list of paragraphs. */
function collectParagraphs(boxEl: HTMLElement): EditedParagraph[] {
	const out: EditedParagraph[] = [];
	for (const block of Array.from(boxEl.children)) {
		if (!(block instanceof HTMLElement)) continue;
		// contenteditable sometimes nests new blocks inside an existing paragraph.
		const nested = Array.from(block.children).filter(
			(el): el is HTMLElement => el instanceof HTMLElement && paragraphRegistry.has(el),
		);
		if (nested.length > 0 && !paragraphRegistry.has(block)) {
			for (const el of nested) out.push(readParagraph(el));
			continue;
		}
		out.push(readParagraph(block));
	}
	return out;
}

function readParagraph(el: HTMLElement): EditedParagraph {
	const runs: { run: Run; text: string }[] = [];
	for (const span of Array.from(el.querySelectorAll<HTMLElement>("[data-run]"))) {
		const run = runRegistry.get(span);
		if (run) runs.push({ run, text: visibleText(span) });
	}
	return {
		model: paragraphRegistry.get(el) ?? null,
		text: visibleText(el),
		runs,
	};
}

/** Text as the user sees it: bullet glyphs excluded, <br> as a newline. */
function visibleText(root: Node): string {
	let out = "";
	const walk = (node: Node): void => {
		if (node.nodeType === Node.TEXT_NODE) {
			out += node.nodeValue ?? "";
			return;
		}
		if (node instanceof HTMLElement) {
			if (node.hasClass("pptx-bullet")) return;
			if (node.tagName === "BR") {
				out += "\n";
				return;
			}
		}
		for (const kid of Array.from(node.childNodes)) walk(kid);
	};
	walk(root);
	return out;
}
