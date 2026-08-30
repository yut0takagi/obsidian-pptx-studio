import { unzipSync, zipSync } from "fflate";
import { attr, parseXml } from "./xml";

export interface Relationship {
	id: string;
	/** Full relationship type URI. */
	type: string;
	/** Last path segment of the type URI, e.g. "slideLayout", "image". */
	kind: string;
	/** Package-absolute target path, e.g. "ppt/slides/slide1.xml". Empty when external. */
	target: string;
	external: boolean;
}

const CONTENT_TYPES = "[Content_Types].xml";
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

const MIME_BY_EXT: Record<string, string> = {
	png: "image/png",
	jpg: "image/jpeg",
	jpeg: "image/jpeg",
	gif: "image/gif",
	bmp: "image/bmp",
	svg: "image/svg+xml",
	webp: "image/webp",
	tif: "image/tiff",
	tiff: "image/tiff",
	avif: "image/avif",
};

/** Raster formats no browser can decode; we show a placeholder instead. */
const UNRENDERABLE = new Set(["emf", "wmf", "tif", "tiff", "eps", "pict", "wdp", "emz", "wmz"]);

/**
 * An unzipped .pptx. Parsed XML parts and image object URLs are memoised, so a
 * layout or theme shared by fifty slides is only parsed once.
 */
export class PptxPackage {
	private readonly files: Record<string, Uint8Array>;
	private readonly xmlCache = new Map<string, Document | null>();
	private readonly relsCache = new Map<string, Map<string, Relationship>>();
	private readonly urlCache = new Map<string, string | null>();
	private readonly objectUrls: string[] = [];
	/** Parts whose parsed Document has been edited and must be re-serialised on save. */
	private readonly dirtyParts = new Set<string>();

	private constructor(files: Record<string, Uint8Array>) {
		this.files = files;
	}

	static open(data: ArrayBuffer | Uint8Array): PptxPackage {
		const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
		if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
			throw new Error("Not a .pptx file (missing ZIP signature).");
		}
		let files: Record<string, Uint8Array>;
		try {
			files = unzipSync(bytes);
		} catch (e) {
			throw new Error(`Could not unzip the .pptx: ${(e as Error).message}`);
		}
		// Normalise away any leading "/" some producers emit.
		const normalised: Record<string, Uint8Array> = {};
		for (const [name, content] of Object.entries(files)) {
			normalised[name.replace(/^\/+/, "")] = content;
		}
		if (!normalised["ppt/presentation.xml"]) {
			throw new Error("Not a PowerPoint presentation (ppt/presentation.xml is missing).");
		}
		return new PptxPackage(normalised);
	}

	has(path: string): boolean {
		return this.files[path] !== undefined;
	}

	bytes(path: string): Uint8Array | null {
		return this.files[path] ?? null;
	}

	/** Parsed XML part, or null when the part is absent or unparseable. */
	xml(path: string): Document | null {
		if (this.xmlCache.has(path)) return this.xmlCache.get(path) ?? null;
		const raw = this.files[path];
		let doc: Document | null = null;
		if (raw) {
			try {
				doc = parseXml(raw);
			} catch {
				doc = null;
			}
		}
		this.xmlCache.set(path, doc);
		return doc;
	}

	/** Relationships declared for a part, keyed by r:id. */
	rels(partPath: string): Map<string, Relationship> {
		const cached = this.relsCache.get(partPath);
		if (cached) return cached;

		const slash = partPath.lastIndexOf("/");
		const dir = slash === -1 ? "" : partPath.slice(0, slash);
		const base = slash === -1 ? partPath : partPath.slice(slash + 1);
		const relsPath = `${dir ? dir + "/" : ""}_rels/${base}.rels`;

		const map = new Map<string, Relationship>();
		const doc = this.xml(relsPath);
		if (doc) {
			for (const el of Array.from(doc.getElementsByTagName("*"))) {
				if (el.localName !== "Relationship") continue;
				const id = attr(el, "Id");
				const type = attr(el, "Type") ?? "";
				const target = attr(el, "Target") ?? "";
				if (!id) continue;
				const external = (attr(el, "TargetMode") ?? "") === "External";
				map.set(id, {
					id,
					type,
					kind: type.slice(type.lastIndexOf("/") + 1),
					target: external ? target : resolvePath(dir, target),
					external,
				});
			}
		}
		this.relsCache.set(partPath, map);
		return map;
	}

	/** Resolve one r:embed / r:id against a part's relationships. */
	relTarget(partPath: string, id: string | null): string | null {
		if (!id) return null;
		const rel = this.rels(partPath).get(id);
		if (!rel || rel.external) return null;
		return rel.target;
	}

	/** First relationship of a given kind, e.g. "slideLayout". */
	relByKind(partPath: string, kind: string): Relationship | null {
		for (const rel of this.rels(partPath).values()) {
			if (rel.kind === kind) return rel;
		}
		return null;
	}

	/**
	 * A displayable object URL for a media part. Returns null for vector formats
	 * browsers cannot decode (EMF/WMF), so callers can draw a placeholder.
	 */
	mediaUrl(path: string | null): string | null {
		if (!path) return null;
		if (this.urlCache.has(path)) return this.urlCache.get(path) ?? null;
		const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
		const raw = this.files[path];
		let url: string | null = null;
		if (raw && !UNRENDERABLE.has(ext)) {
			const mime = MIME_BY_EXT[ext] ?? "application/octet-stream";
			// Copy into a fresh buffer: fflate hands back views onto a shared one.
			url = URL.createObjectURL(new Blob([raw.slice()], { type: mime }));
			this.objectUrls.push(url);
		}
		this.urlCache.set(path, url);
		return url;
	}

	/** Base64 data URL for a media part, used when exporting to PNG. */
	mediaDataUrl(path: string): string | null {
		const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
		const raw = this.files[path];
		if (!raw || UNRENDERABLE.has(ext)) return null;
		const mime = MIME_BY_EXT[ext] ?? "application/octet-stream";
		let binary = "";
		const chunk = 0x8000;
		for (let i = 0; i < raw.length; i += chunk) {
			binary += String.fromCharCode(...raw.subarray(i, i + chunk));
		}
		return `data:${mime};base64,${btoa(binary)}`;
	}

	/**
	 * Mark a part as edited. The part must already have been read through xml(),
	 * because the cached Document is what gets written back.
	 */
	markDirty(path: string): void {
		if (!this.xmlCache.get(path)) {
			throw new Error(`Cannot mark ${path} dirty: it has not been parsed.`);
		}
		this.dirtyParts.add(path);
	}

	get isDirty(): boolean {
		return this.dirtyParts.size > 0;
	}

	get dirtyPartPaths(): string[] {
		return [...this.dirtyParts];
	}

	clearDirty(): void {
		this.dirtyParts.clear();
	}

	/**
	 * Repackage the deck. Edited parts are re-serialised from their Documents;
	 * every other entry is written back byte for byte, so anything this plugin
	 * does not understand — animations, embedded fonts, exotic formatting —
	 * survives the round trip untouched.
	 */
	toZip(): Uint8Array {
		const encoder = new TextEncoder();
		const out: Record<string, Uint8Array> = {};

		// [Content_Types].xml must be the first entry of an OPC package.
		const ordered = Object.keys(this.files).sort((a, b) =>
			a === CONTENT_TYPES ? -1 : b === CONTENT_TYPES ? 1 : 0,
		);

		for (const path of ordered) {
			if (this.dirtyParts.has(path)) {
				const doc = this.xmlCache.get(path);
				if (doc) {
					// Browsers drop the XML declaration when serialising; other DOM
					// implementations keep it. Strip whatever is there and write our own.
					const body = new XMLSerializer()
						.serializeToString(doc)
						.replace(/^\s*<\?xml[^?]*\?>\s*/, "");
					out[path] = encoder.encode(XML_DECL + body);
					continue;
				}
			}
			out[path] = this.files[path];
		}

		return zipSync(out, { level: 6 });
	}

	/** Release every object URL handed out by mediaUrl(). */
	dispose(): void {
		this.dirtyParts.clear();
		for (const url of this.objectUrls) URL.revokeObjectURL(url);
		this.objectUrls.length = 0;
		this.urlCache.clear();
		this.xmlCache.clear();
		this.relsCache.clear();
	}
}

/** Resolve an OPC relationship target (possibly "../media/x.png") against a directory. */
function resolvePath(dir: string, target: string): string {
	if (target.startsWith("/")) return target.slice(1);
	const segments = dir ? dir.split("/") : [];
	for (const part of target.split("/")) {
		if (part === "" || part === ".") continue;
		if (part === "..") segments.pop();
		else segments.push(part);
	}
	return segments.join("/");
}
