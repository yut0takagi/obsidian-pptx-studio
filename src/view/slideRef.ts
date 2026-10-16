import type { Deck } from "../pptx/types";

/**
 * Which slide an embed's subpath names.
 *
 * `#3` (and `#slide3`, `#slide=3`) is a position, which is what embeds have
 * always used and what existing notes contain. `#id=257` is the slide's
 * `p:sldId@id`, which PowerPoint keeps with the slide when the deck is
 * reordered — so an embed written that way keeps showing the same slide after
 * slides are inserted or moved in front of it.
 */
export type SlideRef = { kind: "number"; number: number } | { kind: "id"; id: number };

export function parseSlideRef(subpath: string): SlideRef | undefined {
	const trimmed = subpath.trim();
	const byId = /^id=(\d+)$/i.exec(trimmed);
	if (byId) return { kind: "id", id: Number(byId[1]) };
	const byNumber = /^(?:slide=?)?(\d+)$/i.exec(trimmed);
	if (!byNumber) return undefined;
	const n = Number(byNumber[1]);
	return Number.isFinite(n) && n > 0 ? { kind: "number", number: n } : undefined;
}

/**
 * The 1-based slide number a reference points at in this deck: undefined for
 * no reference (the whole deck), null for an id the deck no longer has.
 */
export function resolveSlideRef(deck: Deck, ref: SlideRef | undefined): number | null | undefined {
	if (!ref) return undefined;
	if (ref.kind === "number") return ref.number;
	const index = deck.slides.findIndex((slide) => slide.slideId === ref.id);
	return index === -1 ? null : index + 1;
}

/** The embed for one slide, by id where it has one. `index` is 1-based. */
export function slideEmbedLink(linktext: string, slide: { slideId: number | null; index: number }): string {
	const subpath = slide.slideId !== null ? `id=${slide.slideId}` : String(slide.index);
	return `![[${linktext}#${subpath}]]`;
}
