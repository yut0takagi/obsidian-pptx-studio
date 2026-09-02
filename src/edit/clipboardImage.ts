/**
 * Taking an image off the system clipboard.
 *
 * A clipboard rarely holds one thing. Copying from a browser or a screenshot
 * tool typically offers the same picture two or three ways at once, and which
 * one arrives first is not something to rely on — so the representation is
 * chosen deliberately rather than by taking whatever came first.
 *
 * Nothing here reads the bytes: that is asynchronous and belongs to the caller.
 * What is decided here is which item to read and what to call it.
 */

/** Clipboard MIME types PowerPoint can store as-is, and the extension to use. */
const EXTENSION_BY_TYPE: Record<string, string> = {
	"image/png": "png",
	"image/jpeg": "jpg",
	"image/jpg": "jpg",
	"image/gif": "gif",
	"image/bmp": "bmp",
	"image/webp": "webp",
	"image/svg+xml": "svg",
};

/**
 * Preference order. PNG first because it is lossless and the one every
 * clipboard produces; SVG last because a deck full of them renders unevenly
 * outside this plugin.
 */
const PREFERENCE = ["png", "jpg", "gif", "webp", "bmp", "svg"];

/** The extension for a clipboard type, ignoring any parameters after a ";". */
export function extensionForImageType(type: string): string | null {
	const bare = type.toLowerCase().split(";")[0].trim();
	return EXTENSION_BY_TYPE[bare] ?? null;
}

export interface ChosenImage<T> {
	file: T;
	extension: string;
}

/** The best image among the clipboard's representations, or null if none is one. */
export function chooseImage<T extends { type: string }>(
	files: readonly T[],
): ChosenImage<T> | null {
	let best: { file: T; extension: string; rank: number } | null = null;
	for (const file of files) {
		const extension = extensionForImageType(file.type);
		if (!extension) continue;
		const rank = PREFERENCE.indexOf(extension);
		if (!best || rank < best.rank) best = { file, extension, rank };
	}
	return best ? { file: best.file, extension: best.extension } : null;
}

/**
 * What to call the pasted image.
 *
 * A screenshot arrives with no name at all, and one dragged out of a browser
 * often arrives with a name that says nothing about its format — so a missing
 * or extension-less name gets one rather than reaching the deck bare.
 */
export function clipboardImageName(name: string | undefined, extension: string): string {
	const trimmed = (name ?? "").trim();
	if (trimmed === "") return `clipboard.${extension}`;
	return /\.[a-z0-9]+$/i.test(trimmed) ? trimmed : `${trimmed}.${extension}`;
}
