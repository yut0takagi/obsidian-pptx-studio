/**
 * The smallest .pptx the parser will open, built in memory.
 *
 * The generated fixture under tests/fixtures is a whole deck — masters, themes,
 * charts, pictures — which is right for the smoke test and far too much to
 * read when a unit test is about one attribute on one slide. This builds just
 * the parts a test names, so what the test is about is the only thing in it.
 */
import { strToU8, zipSync } from "fflate";
import { PptxPackage } from "../../src/pptx/package";

export const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
export const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
export const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

export interface MiniSlide {
	/** The p:sld element's own attributes, e.g. `show="0"`. */
	attrs?: string;
	/** Children of p:sld after p:cSld, e.g. a p:transition. */
	after?: string;
	/** Shapes inside the slide's spTree. */
	shapes?: string;
	/** Speaker notes text; adds a notes slide related to this one. */
	notes?: string;
	/** p:sldId@id; defaults to 256 + position. */
	id?: number;
}

/** One rels document from (id, type, target) triples. */
export function rels(entries: [string, string, string][]): string {
	return (
		`${DECL}<Relationships xmlns="${PKG_REL_NS}">` +
		entries.map(([id, type, target]) => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`).join("") +
		"</Relationships>"
	);
}

function spTree(shapes: string): string {
	return (
		'<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
		"<p:grpSpPr/>" +
		shapes +
		"</p:spTree>"
	);
}

function textShape(id: number, text: string, ph = ""): string {
	return (
		`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
		'<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>' +
		`<p:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`
	);
}

/** Slide XML for one MiniSlide. */
export function slideXml(slide: MiniSlide, n: number): string {
	return (
		`${DECL}<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"${slide.attrs ? ` ${slide.attrs}` : ""}>` +
		`<p:cSld>${spTree(slide.shapes ?? textShape(2, `Slide ${n}`))}</p:cSld>` +
		"<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>" +
		(slide.after ?? "") +
		"</p:sld>"
	);
}

/** The parts of a deck, as text, before zipping — so a test can add its own. */
export function miniDeckParts(slides: MiniSlide[]): Record<string, string> {
	const parts: Record<string, string> = {};
	const overrides: string[] = [
		'<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
	];
	const presRels: [string, string, string][] = [];
	const sldIds: string[] = [];

	slides.forEach((slide, i) => {
		const n = i + 1;
		const path = `ppt/slides/slide${n}.xml`;
		parts[path] = slideXml(slide, n);
		overrides.push(
			`<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
		);
		presRels.push([`rId${n}`, "slide", `slides/slide${n}.xml`]);
		sldIds.push(`<p:sldId id="${slide.id ?? 256 + n}" r:id="rId${n}"/>`);

		const slideRels: [string, string, string][] = [];
		if (slide.notes !== undefined) {
			const notesPath = `ppt/notesSlides/notesSlide${n}.xml`;
			parts[notesPath] =
				`${DECL}<p:notes xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
				`<p:cSld>${spTree(textShape(2, slide.notes, '<p:ph type="body" idx="1"/>'))}</p:cSld>` +
				"<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>";
			parts[`ppt/notesSlides/_rels/notesSlide${n}.xml.rels`] = rels([
				["rId1", "notesMaster", "../notesMasters/notesMaster1.xml"],
				["rId2", "slide", `../slides/slide${n}.xml`],
			]);
			overrides.push(
				`<Override PartName="/${notesPath}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`,
			);
			slideRels.push(["rId2", "notesSlide", `../notesSlides/notesSlide${n}.xml`]);
		}
		parts[`ppt/slides/_rels/slide${n}.xml.rels`] = rels(slideRels);
	});

	if (slides.some((s) => s.notes !== undefined)) {
		parts["ppt/notesMasters/notesMaster1.xml"] =
			`${DECL}<p:notesMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
			`<p:cSld>${spTree("")}</p:cSld></p:notesMaster>`;
		overrides.push(
			'<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>',
		);
	}

	parts["[Content_Types].xml"] =
		`${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
		'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
		'<Default Extension="xml" ContentType="application/xml"/>' +
		overrides.join("") +
		"</Types>";
	parts["_rels/.rels"] = rels([["rId1", "officeDocument", "ppt/presentation.xml"]]);
	parts["ppt/presentation.xml"] =
		`${DECL}<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
		`<p:sldIdLst>${sldIds.join("")}</p:sldIdLst>` +
		'<p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>';
	parts["ppt/_rels/presentation.xml.rels"] = rels(presRels);
	return parts;
}

/** Zip parts into an open package. */
export function packageOf(parts: Record<string, string>): PptxPackage {
	const files: Record<string, Uint8Array> = {};
	for (const [path, text] of Object.entries(parts)) files[path] = strToU8(text);
	return PptxPackage.open(zipSync(files));
}

export function miniDeck(slides: MiniSlide[]): PptxPackage {
	return packageOf(miniDeckParts(slides));
}

/** A part's current text, edits included. */
export function partText(pkg: PptxPackage, path: string): string | null {
	const bytes = pkg.serializePart(path);
	return bytes ? new TextDecoder().decode(bytes) : null;
}
