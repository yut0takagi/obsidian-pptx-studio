/**
 * Find and replace across a whole deck.
 *
 * The hard part is that a phrase a reader sees as one word is rarely one run:
 * PowerPoint splits a paragraph at every change of formatting, and at every
 * point the author paused typing. Searching run by run would miss "quarterly"
 * whenever the "ly" happened to be bold, so a paragraph is matched as the text
 * it reads as, and a replacement is written back across however many runs the
 * match turned out to span.
 *
 * The replacement takes the formatting of the run the match starts in, which is
 * the same bargain the text editor strikes when an edit crosses a run boundary.
 *
 * Everything here reads the XML rather than the parsed model, so what is
 * searched and what is rewritten can never disagree.
 */
import type { Deck, Shape, Slide } from "../pptx/types";
import { child } from "../pptx/xml";
import type { CommandContext } from "./commands";
import { setRunText } from "./textEdit";

export interface FindOptions {
	matchCase: boolean;
}

export interface TextMatch {
	/** 0-based, matching `deck.slides`. */
	slideIndex: number;
	partPath: string;
	shapeId: string;
	shapeName: string;
	/** The paragraph the hit sits in, for showing it in context. */
	text: string;
	start: number;
	length: number;
	/**
	 * False when the hit touches something that cannot carry typed text — a line
	 * break, or a field like a slide number, whose text the deck regenerates.
	 */
	replaceable: boolean;
}

/** A run as the paragraph reads it, and whether its text is ours to rewrite. */
export interface RunText {
	text: string;
	editable: boolean;
}

export interface Span {
	start: number;
	length: number;
}

// ------------------------------------------------------------- text matching

/** Every non-overlapping occurrence, left to right. A literal search, not a regex. */
export function findInText(text: string, query: string, matchCase: boolean): Span[] {
	if (query === "") return [];
	const haystack = matchCase ? text : text.toLowerCase();
	const needle = matchCase ? query : query.toLowerCase();
	const out: Span[] = [];
	let at = haystack.indexOf(needle);
	while (at !== -1) {
		out.push({ start: at, length: needle.length });
		at = haystack.indexOf(needle, at + needle.length);
	}
	return out;
}

interface Placed extends Span {
	startRun: number;
	startOffset: number;
	endRun: number;
	endOffset: number;
}

/** Locate a span against the run boundaries of the paragraph it came from. */
function place(runs: RunText[], span: Span): Placed | null {
	const end = span.start + span.length;
	let startRun = -1;
	let startOffset = 0;
	let endRun = -1;
	let endOffset = 0;
	let at = 0;
	for (let i = 0; i < runs.length; i++) {
		const next = at + runs[i].text.length;
		if (startRun === -1 && span.start < next) {
			startRun = i;
			startOffset = span.start - at;
		}
		if (endRun === -1 && end <= next) {
			endRun = i;
			endOffset = end - at;
			break;
		}
		at = next;
	}
	if (startRun === -1 || endRun === -1) return null;
	return { ...span, startRun, startOffset, endRun, endOffset };
}

/** True when every run the span touches is one we may rewrite. */
function isReplaceable(runs: RunText[], placed: Placed): boolean {
	for (let i = placed.startRun; i <= placed.endRun; i++) {
		if (!runs[i].editable) return false;
	}
	return true;
}

/**
 * The text each run should end up holding once every span is replaced.
 *
 * Returns null when nothing would change, so a caller can skip the paragraph
 * without comparing strings itself.
 */
export function planReplacement(
	runs: RunText[],
	spans: Span[],
	replacement: string,
): string[] | null {
	const placed = spans
		.map((span) => place(runs, span))
		.filter((p): p is Placed => p !== null && isReplaceable(runs, p))
		.sort((a, b) => a.start - b.start);
	if (placed.length === 0) return null;

	const out = runs.map((run) => run.text);
	for (let i = 0; i < runs.length; i++) {
		const original = runs[i].text;
		let built = "";
		let cursor = 0;
		for (const span of placed) {
			if (span.startRun > i || span.endRun < i) continue;
			if (span.startRun === i) {
				// The replacement lands whole in the run the match started in.
				built += original.slice(cursor, span.startOffset) + replacement;
				cursor = span.endRun === i ? span.endOffset : original.length;
			} else if (span.endRun === i) {
				cursor = Math.max(cursor, span.endOffset);
			} else {
				cursor = original.length;
			}
		}
		out[i] = built + original.slice(cursor);
	}
	return out.every((text, i) => text === runs[i].text) ? null : out;
}

// -------------------------------------------------------------- reading a deck

interface BodyRef {
	shape: Shape;
	/** The a:txBody element. */
	source: Element;
}

/** Every text body a user could type into, including inside groups and tables. */
function editableBodies(shape: Shape, out: BodyRef[]): void {
	switch (shape.kind) {
		case "shape":
			if (shape.text?.source) out.push({ shape, source: shape.text.source });
			break;
		case "table":
			for (const row of shape.table.rows) {
				for (const cell of row.cells) {
					if (!cell.merged && cell.text?.source) out.push({ shape, source: cell.text.source });
				}
			}
			break;
		case "group":
			for (const kid of shape.children) editableBodies(kid, out);
			break;
		default:
			break;
	}
}

/** The shapes a slide owns, skipping the ones drawn from its layout or master. */
function ownBodies(slide: Slide): BodyRef[] {
	const out: BodyRef[] = [];
	for (const shape of slide.shapes.slice(slide.templateShapes)) editableBodies(shape, out);
	return out;
}

/** Paragraph elements of a text body, in order. */
function paragraphsOf(body: Element): Element[] {
	const out: Element[] = [];
	for (let n = body.firstElementChild; n; n = n.nextElementSibling) {
		if (n.localName === "p") out.push(n);
	}
	return out;
}

/**
 * A paragraph as a list of runs, paired with the elements to write back to.
 *
 * `a:br` reads as the newline it draws and `a:fld` as its cached text, so a
 * search sees the paragraph the way the slide does — but neither is ours to
 * rewrite, so a match touching one is reported and left alone.
 */
function runsOf(paragraph: Element): { runs: RunText[]; elements: Element[] } {
	const runs: RunText[] = [];
	const elements: Element[] = [];
	for (let n = paragraph.firstElementChild; n; n = n.nextElementSibling) {
		if (n.localName === "r") {
			runs.push({ text: child(n, "t")?.textContent ?? "", editable: true });
			elements.push(n);
		} else if (n.localName === "fld") {
			runs.push({ text: child(n, "t")?.textContent ?? "", editable: false });
			elements.push(n);
		} else if (n.localName === "br") {
			runs.push({ text: "\n", editable: false });
			elements.push(n);
		}
	}
	return { runs, elements };
}

export function findMatches(deck: Deck, query: string, options: FindOptions): TextMatch[] {
	const out: TextMatch[] = [];
	if (query === "") return out;
	deck.slides.forEach((slide, slideIndex) => {
		for (const { shape, source } of ownBodies(slide)) {
			for (const paragraph of paragraphsOf(source)) {
				const { runs } = runsOf(paragraph);
				const text = runs.map((r) => r.text).join("");
				for (const span of findInText(text, query, options.matchCase)) {
					const placed = place(runs, span);
					out.push({
						slideIndex,
						partPath: slide.partPath,
						shapeId: shape.id,
						shapeName: shape.name,
						text,
						start: span.start,
						length: span.length,
						replaceable: placed !== null && isReplaceable(runs, placed),
					});
				}
			}
		}
	});
	return out;
}

/**
 * Replace every occurrence in the deck, as one undoable edit.
 *
 * Returns how many were replaced — which can be fewer than were found, when a
 * hit runs across a line break or a field.
 */
export function replaceAll(
	ctx: CommandContext,
	query: string,
	replacement: string,
	options: FindOptions,
	label: string,
): number {
	if (query === "") return 0;
	const parts = [
		...new Set(
			findMatches(ctx.deck, query, options)
				.filter((m) => m.replaceable)
				.map((m) => m.partPath),
		),
	];
	if (parts.length === 0) return 0;

	let replaced = 0;
	const done = ctx.editor.transact(label, parts, () => {
		for (const slide of ctx.deck.slides) {
			if (!parts.includes(slide.partPath)) continue;
			for (const { source } of ownBodies(slide)) {
				for (const paragraph of paragraphsOf(source)) {
					const { runs, elements } = runsOf(paragraph);
					const text = runs.map((r) => r.text).join("");
					const spans = findInText(text, query, options.matchCase);
					if (spans.length === 0) continue;
					const planned = planReplacement(runs, spans, replacement);
					if (!planned) continue;
					planned.forEach((value, i) => {
						if (value === runs[i].text || !runs[i].editable) return;
						setRunText(elements[i], value);
					});
					replaced += spans.filter((span) => {
						const placed = place(runs, span);
						return placed !== null && isReplaceable(runs, placed);
					}).length;
				}
			}
		}
		return replaced > 0;
	});
	return done ? replaced : 0;
}
