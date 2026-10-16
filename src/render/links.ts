/**
 * Which hyperlink targets in a slide become links.
 *
 * A link's target comes from the deck's relationships, and a deck may come
 * from anyone. Web and mail links are what slides link to; a script URL, a
 * local file, or an app's own scheme is not something a click on slide text
 * should be able to start, so those stay plain text.
 */
const ALLOWED = new Set(["http:", "https:", "mailto:"]);

/** The href for a run's link, or null to draw it as text. */
export function safeHref(link: string | null): string | null {
	if (!link) return null;
	let url: URL;
	try {
		// Parsing the way the browser will — dropping stray whitespace, folding
		// case — is what keeps "java\tscript:" from slipping past a string check.
		url = new URL(link);
	} catch {
		// Relative targets (another slide, a part in the package) land here too.
		return null;
	}
	return ALLOWED.has(url.protocol) ? url.href : null;
}
