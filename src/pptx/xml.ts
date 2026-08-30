/**
 * Minimal XML helpers for OOXML parsing.
 *
 * Everything matches on *local* names rather than the `a:` / `p:` prefixes that
 * PowerPoint happens to emit, so a deck written by a different producer with
 * different prefix bindings still parses.
 */

const decoder = new TextDecoder("utf-8");

export function parseXml(bytes: Uint8Array): Document {
	let text = decoder.decode(bytes);
	if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
	const doc = new DOMParser().parseFromString(text, "application/xml");
	// Browsers report failures as a <parsererror> element rather than by throwing.
	const err = doc.getElementsByTagName("parsererror")[0];
	if (err) throw new Error(`Malformed XML: ${err.textContent?.slice(0, 200)}`);
	return doc;
}

/** Direct children with the given local name. */
export function children(el: Element | null | undefined, localName: string): Element[] {
	if (!el) return [];
	const out: Element[] = [];
	for (let n = el.firstElementChild; n; n = n.nextElementSibling) {
		if (n.localName === localName) out.push(n);
	}
	return out;
}

/** First direct child with the given local name. */
export function child(el: Element | null | undefined, localName: string): Element | null {
	if (!el) return null;
	for (let n = el.firstElementChild; n; n = n.nextElementSibling) {
		if (n.localName === localName) return n;
	}
	return null;
}

/** Walk a chain of direct children, e.g. childPath(sp, "spPr", "xfrm", "off"). */
export function childPath(el: Element | null | undefined, ...path: string[]): Element | null {
	let cur: Element | null = el ?? null;
	for (const name of path) {
		cur = child(cur, name);
		if (!cur) return null;
	}
	return cur;
}

/** All descendants with the given local name, in document order. */
export function descendants(el: Element | null | undefined, localName: string): Element[] {
	if (!el) return [];
	const out: Element[] = [];
	const walk = (node: Element) => {
		for (let n = node.firstElementChild; n; n = n.nextElementSibling) {
			if (n.localName === localName) out.push(n);
			walk(n);
		}
	};
	walk(el);
	return out;
}

/** First descendant with the given local name. */
export function descendant(el: Element | null | undefined, localName: string): Element | null {
	if (!el) return null;
	for (let n = el.firstElementChild; n; n = n.nextElementSibling) {
		if (n.localName === localName) return n;
		const deep = descendant(n, localName);
		if (deep) return deep;
	}
	return null;
}

/** Attribute value, ignoring any namespace prefix on the attribute name. */
export function attr(el: Element | null | undefined, name: string): string | null {
	if (!el) return null;
	const direct = el.getAttribute(name);
	if (direct !== null) return direct;
	for (const a of Array.from(el.attributes)) {
		if (a.localName === name) return a.value;
	}
	return null;
}

export function numAttr(el: Element | null | undefined, name: string): number | null {
	const raw = attr(el, name);
	if (raw === null || raw === "") return null;
	const n = Number(raw);
	return Number.isFinite(n) ? n : null;
}

/** OOXML booleans are "1"/"0"/"true"/"false"; a present-but-empty attr means true. */
export function boolAttr(el: Element | null | undefined, name: string): boolean | null {
	const raw = attr(el, name);
	if (raw === null) return null;
	return raw === "1" || raw === "true" || raw === "";
}
