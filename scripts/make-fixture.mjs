/**
 * Build tests/fixtures/sample.pptx.
 *
 * A hand-written deck is a better fixture than a real one: it is small, it is
 * reproducible, and every feature the renderer claims to support appears in it
 * exactly once, so a regression is obvious on sight.
 *
 *   node scripts/make-fixture.mjs [outfile]
 */
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import { zipSync } from "fflate";
import { encodePng } from "./png.mjs";

/** Slides are authored in pixels; OOXML wants English Metric Units. */
const px = (n) => Math.round(n * 9525);
const SLIDE_W = 1280;
const SLIDE_H = 720;

const NS =
	'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
	'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
	'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

const esc = (s) =>
	String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

// ----------------------------------------------------------- drawing helpers

function xfrm(x, y, w, h, rot) {
	const rotAttr = rot ? ` rot="${Math.round(rot * 60000)}"` : "";
	return `<a:xfrm${rotAttr}><a:off x="${px(x)}" y="${px(y)}"/><a:ext cx="${px(w)}" cy="${px(h)}"/></a:xfrm>`;
}

function fillXml(fill) {
	if (!fill) return "";
	if (fill === "none") return "<a:noFill/>";
	if (fill.solid) return `<a:solidFill><a:srgbClr val="${fill.solid}"/></a:solidFill>`;
	if (fill.scheme) {
		const mods = (fill.lumMod ? `<a:lumMod val="${fill.lumMod * 1000}"/>` : "") +
			(fill.lumOff ? `<a:lumOff val="${fill.lumOff * 1000}"/>` : "");
		return `<a:solidFill><a:schemeClr val="${fill.scheme}">${mods}</a:schemeClr></a:solidFill>`;
	}
	if (fill.gradient) {
		const stops = fill.gradient
			.map(([pos, color]) => `<a:gs pos="${pos * 1000}"><a:srgbClr val="${color}"/></a:gs>`)
			.join("");
		return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:lin ang="${
			Math.round((fill.angle ?? 45) * 60000)
		}" scaled="0"/></a:gradFill>`;
	}
	if (fill.image) return `<a:blipFill><a:blip r:embed="${fill.image}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`;
	return "";
}

function lineXml(line) {
	if (!line) return "";
	const dash = line.dash ? `<a:prstDash val="${line.dash}"/>` : "";
	return `<a:ln w="${px(line.w ?? 1)}">${fillXml(line.fill ?? { solid: "000000" })}${dash}</a:ln>`;
}

function runXml(run) {
	if (run.br) return "<a:br/>";
	const props = [
		`lang="en-US"`,
		run.sz ? `sz="${Math.round(run.sz * 100)}"` : "",
		run.b ? 'b="1"' : "",
		run.i ? 'i="1"' : "",
		run.u ? 'u="sng"' : "",
		run.strike ? 'strike="sngStrike"' : "",
	]
		.filter(Boolean)
		.join(" ");
	const color = run.color ? `<a:solidFill><a:srgbClr val="${run.color}"/></a:solidFill>` : "";
	const font = run.font ? `<a:latin typeface="${esc(run.font)}"/>` : "";
	const link = run.link ? `<a:hlinkClick r:id="${run.link}"/>` : "";
	return `<a:r><a:rPr ${props}>${color}${font}${link}</a:rPr><a:t>${esc(run.t)}</a:t></a:r>`;
}

function paragraphXml(para) {
	const props = [
		para.lvl ? `lvl="${para.lvl}"` : "",
		para.algn ? `algn="${para.algn}"` : "",
		para.marL !== undefined ? `marL="${px(para.marL)}"` : "",
		para.indent !== undefined ? `indent="${px(para.indent)}"` : "",
	]
		.filter(Boolean)
		.join(" ");
	let bullet = "";
	if (para.bullet === "none") bullet = "<a:buNone/>";
	else if (para.bullet === "number") bullet = '<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>';
	else if (para.bullet) bullet = `<a:buChar char="${esc(para.bullet)}"/>`;
	const pPr = props || bullet ? `<a:pPr ${props}>${bullet}</a:pPr>` : "";
	const runs = (para.runs ?? []).map(runXml).join("");
	return `<a:p>${pPr}${runs}</a:p>`;
}

function textBodyXml(text) {
	if (!text) return "";
	const anchor = text.anchor ? ` anchor="${text.anchor}"` : "";
	const wrap = text.wrap === false ? ' wrap="none"' : "";
	return (
		`<p:txBody><a:bodyPr${anchor}${wrap}/><a:lstStyle/>` +
		text.paragraphs.map(paragraphXml).join("") +
		"</p:txBody>"
	);
}

let nextId = 100;
const id = () => nextId++;

function shape({ name, x, y, w, h, geom = "rect", rot, fill, line, text, ph }) {
	const phXml = ph ? `<p:ph type="${ph.type}"${ph.idx !== undefined ? ` idx="${ph.idx}"` : ""}/>` : "";
	const geometry = `<a:prstGeom prst="${geom}"><a:avLst/></a:prstGeom>`;
	const position = x === undefined ? "" : xfrm(x, y, w, h, rot);
	return (
		`<p:sp><p:nvSpPr><p:cNvPr id="${id()}" name="${esc(name)}"/><p:cNvSpPr/>` +
		`<p:nvPr>${phXml}</p:nvPr></p:nvSpPr>` +
		`<p:spPr>${position}${geometry}${fillXml(fill)}${lineXml(line)}</p:spPr>` +
		textBodyXml(text) +
		"</p:sp>"
	);
}

function picture({ name, x, y, w, h, embed, crop }) {
	const srcRect = crop
		? `<a:srcRect l="${crop.l * 1000}" t="${crop.t * 1000}" r="${crop.r * 1000}" b="${crop.b * 1000}"/>`
		: "";
	return (
		`<p:pic><p:nvPicPr><p:cNvPr id="${id()}" name="${esc(name)}" descr="${esc(name)}"/>` +
		`<p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
		`<p:blipFill><a:blip r:embed="${embed}"/>${srcRect}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
		`<p:spPr>${xfrm(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
	);
}

function connector({ name, x, y, w, h, line }) {
	return (
		`<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id()}" name="${esc(name)}"/><p:cNvCxnSpPr/>` +
		`<p:nvPr/></p:nvCxnSpPr>` +
		`<p:spPr>${xfrm(x, y, w, h)}<a:prstGeom prst="line"><a:avLst/></a:prstGeom>${lineXml(line)}</p:spPr>` +
		"</p:cxnSp>"
	);
}

function group({ name, x, y, w, h, childOffset, children }) {
	const co = childOffset ?? { x: 0, y: 0, w, h };
	const frame =
		`<a:xfrm><a:off x="${px(x)}" y="${px(y)}"/><a:ext cx="${px(w)}" cy="${px(h)}"/>` +
		`<a:chOff x="${px(co.x)}" y="${px(co.y)}"/><a:chExt cx="${px(co.w)}" cy="${px(co.h)}"/></a:xfrm>`;
	return (
		`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id()}" name="${esc(name)}"/><p:cNvGrpSpPr/>` +
		`<p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${frame}</p:grpSpPr>${children.join("")}</p:grpSp>`
	);
}

function table({ name, x, y, w, h, columns, rows }) {
	const grid = columns.map((cw) => `<a:gridCol w="${px(cw)}"/>`).join("");
	const body = rows
		.map((row) => {
			const cells = row.cells
				.map((cell) => {
					if (cell.hMerge) return '<a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc>';
					const span = cell.colSpan ? ` gridSpan="${cell.colSpan}"` : "";
					const tcPr =
						`<a:tcPr marL="${px(8)}" marR="${px(8)}" marT="${px(4)}" marB="${px(4)}" anchor="ctr">` +
						`<a:lnL w="${px(0.75)}"><a:solidFill><a:srgbClr val="C9D1DB"/></a:solidFill></a:lnL>` +
						`<a:lnR w="${px(0.75)}"><a:solidFill><a:srgbClr val="C9D1DB"/></a:solidFill></a:lnR>` +
						`<a:lnT w="${px(0.75)}"><a:solidFill><a:srgbClr val="C9D1DB"/></a:solidFill></a:lnT>` +
						`<a:lnB w="${px(0.75)}"><a:solidFill><a:srgbClr val="C9D1DB"/></a:solidFill></a:lnB>` +
						fillXml(cell.fill) +
						"</a:tcPr>";
					const text = textBodyXml({
						paragraphs: [{ algn: cell.algn, runs: [{ t: cell.text, b: cell.b, sz: 14, color: cell.color }] }],
					}).replace(/^<p:txBody>/, "<a:txBody>").replace(/<\/p:txBody>$/, "</a:txBody>");
					return `<a:tc${span}>${text}${tcPr}</a:tc>`;
				})
				.join("");
			return `<a:tr h="${px(row.height)}">${cells}</a:tr>`;
		})
		.join("");
	return (
		`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id()}" name="${esc(name)}"/>` +
		`<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
		`<p:xfrm><a:off x="${px(x)}" y="${px(y)}"/><a:ext cx="${px(w)}" cy="${px(h)}"/></p:xfrm>` +
		`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
		`<a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>${grid}</a:tblGrid>${body}</a:tbl>` +
		"</a:graphicData></a:graphic></p:graphicFrame>"
	);
}

function spTree(shapes) {
	return (
		`<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
		`<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>` +
		`<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
		shapes.join("") +
		"</p:spTree>"
	);
}

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

function slideXml(shapes, { background } = {}) {
	const bg = background ? `<p:bg><p:bgPr>${fillXml(background)}<a:effectLst/></p:bgPr></p:bg>` : "";
	return `${DECL}<p:sld ${NS}><p:cSld>${bg}${spTree(shapes)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

function rels(entries) {
	const items = entries
		.map(
			(e) =>
				`<Relationship Id="${e.id}" Type="http://schemas.openxmlformats.org/${e.type}" Target="${esc(
					e.target,
				)}"${e.external ? ' TargetMode="External"' : ""}/>`,
		)
		.join("");
	return `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
}

// -------------------------------------------------------------------- theme

const THEME = `${DECL}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Fixture">
<a:themeElements>
<a:clrScheme name="Fixture">
<a:dk1><a:srgbClr val="1B2733"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="2E4057"/></a:dk2><a:lt2><a:srgbClr val="EEF2F6"/></a:lt2>
<a:accent1><a:srgbClr val="2F6FED"/></a:accent1><a:accent2><a:srgbClr val="E8590C"/></a:accent2>
<a:accent3><a:srgbClr val="12B886"/></a:accent3><a:accent4><a:srgbClr val="F59F00"/></a:accent4>
<a:accent5><a:srgbClr val="7048E8"/></a:accent5><a:accent6><a:srgbClr val="E64980"/></a:accent6>
<a:hlink><a:srgbClr val="1971C2"/></a:hlink><a:folHlink><a:srgbClr val="9C36B5"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Fixture">
<a:majorFont><a:latin typeface="Helvetica Neue"/><a:ea typeface="Hiragino Sans"/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Helvetica Neue"/><a:ea typeface="Hiragino Sans"/><a:cs typeface=""/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Fixture">
<a:fillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"><a:shade val="80000"/></a:schemeClr></a:solidFill>
</a:fillStyleLst>
<a:lnStyleLst>
<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
<a:ln w="28575"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
</a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"><a:shade val="90000"/></a:schemeClr></a:solidFill>
</a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
</a:theme>`;

// ------------------------------------------------------- master and layouts

function levelStyle(level, size, bold, bullet) {
	const buttons =
		bullet === "none"
			? "<a:buNone/>"
			: `<a:buFont typeface="Arial"/><a:buChar char="${bullet}"/>`;
	return (
		`<a:lvl${level}pPr marL="${px(level === 1 ? 0 : (level - 1) * 28)}" indent="${px(-20)}">` +
		buttons +
		`<a:defRPr sz="${size * 100}"${bold ? ' b="1"' : ""}><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>` +
		'<a:latin typeface="+mn-lt"/></a:defRPr></a:lvl' + level + "pPr>"
	);
}

const MASTER = `${DECL}<p:sldMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>${spTree([
	shape({
		name: "Master title",
		x: 80,
		y: 60,
		w: 1120,
		h: 100,
		fill: "none",
		ph: { type: "title" },
		text: { anchor: "b", paragraphs: [{ runs: [] }] },
	}),
	shape({
		name: "Master body",
		x: 80,
		y: 190,
		w: 1120,
		h: 460,
		fill: "none",
		ph: { type: "body", idx: 1 },
		text: { paragraphs: [{ runs: [] }] },
	}),
])}</p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst>
<p:txStyles>
<p:titleStyle><a:lvl1pPr algn="l"><a:buNone/><a:defRPr sz="4000" b="1"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>
<p:bodyStyle>${levelStyle(1, 22, false, "●")}${levelStyle(2, 19, false, "○")}${levelStyle(3, 17, false, "▪")}${levelStyle(4, 15, false, "–")}</p:bodyStyle>
<p:otherStyle><a:lvl1pPr><a:defRPr sz="1400"/></a:lvl1pPr></p:otherStyle>
</p:txStyles></p:sldMaster>`;

const TITLE_LAYOUT = `${DECL}<p:sldLayout ${NS} type="title" preserve="1"><p:cSld name="Title Slide">${spTree([
	shape({
		name: "Title",
		x: 100,
		y: 250,
		w: 1080,
		h: 140,
		fill: "none",
		ph: { type: "ctrTitle" },
		text: { anchor: "b", paragraphs: [{ runs: [] }] },
	}),
	shape({
		name: "Subtitle",
		x: 100,
		y: 400,
		w: 1080,
		h: 80,
		fill: "none",
		ph: { type: "subTitle", idx: 1 },
		text: {
			paragraphs: [{ bullet: "none", runs: [] }],
		},
	}),
])}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

const CONTENT_LAYOUT = `${DECL}<p:sldLayout ${NS} type="obj" preserve="1"><p:cSld name="Title and Content">${spTree([
	// Layout decoration: not a placeholder, so it is drawn on every slide using
	// this layout. Exercises the layout/master artwork path.
	shape({ name: "Accent bar", x: 0, y: 0, w: 1280, h: 8, fill: { scheme: "accent1" } }),
	shape({
		name: "Footer rule",
		x: 80,
		y: 668,
		w: 1120,
		h: 1,
		fill: { solid: "D6DEE8" },
	}),
	shape({
		name: "Title",
		x: 80,
		y: 48,
		w: 1120,
		h: 80,
		fill: "none",
		ph: { type: "title" },
		text: { anchor: "b", paragraphs: [{ runs: [] }] },
	}),
	shape({
		name: "Content",
		x: 80,
		y: 160,
		w: 1120,
		h: 480,
		fill: "none",
		ph: { type: "body", idx: 1 },
		text: { paragraphs: [{ runs: [] }] },
	}),
])}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

const NOTES_MASTER = `${DECL}<p:notesMaster ${NS}><p:cSld>${spTree([])}</p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:notesStyle><a:lvl1pPr><a:defRPr sz="1200"/></a:lvl1pPr></p:notesStyle></p:notesMaster>`;

function notesSlideXml(text) {
	return `${DECL}<p:notes ${NS}><p:cSld>${spTree([
		shape({
			name: "Notes Placeholder",
			x: 40,
			y: 40,
			w: 600,
			h: 300,
			fill: "none",
			ph: { type: "body", idx: 1 },
			text: { paragraphs: text.map((t) => ({ bullet: "none", runs: [{ t }] })) },
		}),
	])}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}

// ------------------------------------------------------------------ slides

const slide1 = slideXml(
	[
		shape({
			name: "Title 1",
			ph: { type: "ctrTitle" },
			text: {
				anchor: "b",
				paragraphs: [{ runs: [{ t: "PPTX Viewer" }] }],
			},
		}),
		shape({
			name: "Subtitle 2",
			ph: { type: "subTitle", idx: 1 },
			text: {
				paragraphs: [
					{ bullet: "none", runs: [{ t: "A fixture deck — every feature the renderer claims, exactly once", sz: 20, color: "51606F" }] },
				],
			},
		}),
	],
	{ background: { gradient: [[0, "FFFFFF"], [100, "DCE6F5"]], angle: 90 } },
);

const slide2 = slideXml([
	shape({
		name: "Title",
		ph: { type: "title" },
		text: { anchor: "b", paragraphs: [{ runs: [{ t: "Bullets, levels and numbering" }] }] },
	}),
	shape({
		name: "Content",
		ph: { type: "body", idx: 1 },
		text: {
			paragraphs: [
				{ runs: [{ t: "Level one inherits its bullet from the master" }] },
				{ lvl: 1, runs: [{ t: "Level two, indented and smaller" }] },
				{ lvl: 2, runs: [{ t: "Level three keeps going" }] },
				{ bullet: "none", runs: [{ t: "" }] },
				{ bullet: "number", runs: [{ t: "Auto-numbered items restart per level" }] },
				{ bullet: "number", runs: [{ t: "Second numbered item" }] },
				{ bullet: "number", runs: [{ t: "Third numbered item" }] },
				{ bullet: "none", runs: [{ t: "" }] },
				{
					bullet: "none",
					runs: [
						{ t: "Mixed runs: " },
						{ t: "bold", b: true },
						{ t: ", " },
						{ t: "italic", i: true },
						{ t: ", " },
						{ t: "underlined", u: true },
						{ t: ", " },
						{ t: "coloured", color: "E8590C", b: true },
						{ t: " — all in one paragraph." },
					],
				},
			],
		},
	}),
]);

const slide3 = slideXml([
	shape({
		name: "Title",
		ph: { type: "title" },
		text: { anchor: "b", paragraphs: [{ runs: [{ t: "Shapes, fills and outlines" }] }] },
	}),
	shape({
		name: "Rect",
		x: 80,
		y: 190,
		w: 200,
		h: 120,
		fill: { scheme: "accent1" },
		text: { anchor: "ctr", paragraphs: [{ algn: "ctr", bullet: "none", runs: [{ t: "rect", color: "FFFFFF", b: true }] }] },
	}),
	shape({
		name: "RoundRect",
		x: 310,
		y: 190,
		w: 200,
		h: 120,
		geom: "roundRect",
		fill: { gradient: [[0, "12B886"], [100, "0B7285"]], angle: 45 },
		text: { anchor: "ctr", paragraphs: [{ algn: "ctr", bullet: "none", runs: [{ t: "roundRect", color: "FFFFFF", b: true }] }] },
	}),
	shape({
		name: "Ellipse",
		x: 540,
		y: 190,
		w: 200,
		h: 120,
		geom: "ellipse",
		fill: { scheme: "accent4" },
		line: { w: 3, fill: { solid: "8B5A00" } },
		text: { anchor: "ctr", paragraphs: [{ algn: "ctr", bullet: "none", runs: [{ t: "ellipse", b: true }] }] },
	}),
	shape({ name: "Triangle", x: 770, y: 190, w: 140, h: 120, geom: "triangle", fill: { scheme: "accent5" } }),
	shape({ name: "Star", x: 930, y: 190, w: 120, h: 120, geom: "star5", fill: { scheme: "accent6" } }),
	shape({
		name: "Arrow",
		x: 1070,
		y: 205,
		w: 130,
		h: 90,
		geom: "rightArrow",
		fill: { scheme: "accent2" },
	}),
	shape({
		name: "Rotated",
		x: 80,
		y: 380,
		w: 220,
		h: 90,
		rot: -8,
		fill: { solid: "FFFFFF" },
		line: { w: 2, fill: { scheme: "accent1" }, dash: "dash" },
		text: { anchor: "ctr", paragraphs: [{ algn: "ctr", bullet: "none", runs: [{ t: "rotated −8°, dashed" }] }] },
	}),
	group({
		name: "Group",
		x: 340,
		y: 370,
		w: 400,
		h: 110,
		childOffset: { x: 0, y: 0, w: 400, h: 110 },
		children: [
			shape({ name: "G1", x: 0, y: 0, w: 120, h: 110, geom: "roundRect", fill: { scheme: "accent3" } }),
			shape({ name: "G2", x: 140, y: 0, w: 120, h: 110, geom: "roundRect", fill: { scheme: "accent3", lumMod: 60, lumOff: 40 } }),
			shape({ name: "G3", x: 280, y: 0, w: 120, h: 110, geom: "roundRect", fill: { scheme: "accent3", lumMod: 40, lumOff: 60 } }),
		],
	}),
	connector({ name: "Connector", x: 790, y: 380, w: 300, h: 90, line: { w: 2.5, fill: { solid: "2E4057" } } }),
	shape({
		name: "Caption",
		x: 80,
		y: 510,
		w: 1120,
		h: 40,
		fill: "none",
		text: {
			paragraphs: [
				{
					bullet: "none",
					runs: [{ t: "Group children scale with the group; the connector is drawn as SVG.", sz: 15, color: "6B7B8C", i: true }],
				},
			],
		},
	}),
]);

const slide4 = slideXml([
	shape({
		name: "Title",
		ph: { type: "title" },
		text: { anchor: "b", paragraphs: [{ runs: [{ t: "Tables" }] }] },
	}),
	table({
		name: "Table",
		x: 120,
		y: 200,
		w: 1040,
		h: 260,
		columns: [340, 240, 240, 220],
		rows: [
			{
				height: 52,
				cells: [
					{ text: "Feature", b: true, color: "FFFFFF", fill: { solid: "2F6FED" } },
					{ text: "Viewer", b: true, color: "FFFFFF", fill: { solid: "2F6FED" }, algn: "ctr" },
					{ text: "Embed", b: true, color: "FFFFFF", fill: { solid: "2F6FED" }, algn: "ctr" },
					{ text: "Export", b: true, color: "FFFFFF", fill: { solid: "2F6FED" }, algn: "ctr" },
				],
			},
			{
				height: 46,
				cells: [
					{ text: "Text and bullets" },
					{ text: "yes", algn: "ctr" },
					{ text: "yes", algn: "ctr" },
					{ text: "yes", algn: "ctr" },
				],
			},
			{
				height: 46,
				cells: [
					{ text: "Editing", fill: { solid: "F2F6FC" } },
					{ text: "yes", algn: "ctr", fill: { solid: "F2F6FC" } },
					{ text: "read-only", algn: "ctr", fill: { solid: "F2F6FC" } },
					{ text: "n/a", algn: "ctr", fill: { solid: "F2F6FC" } },
				],
			},
			{
				height: 46,
				cells: [
					{ text: "Merged cell spanning three columns", colSpan: 3, b: true },
					{ hMerge: true },
					{ hMerge: true },
					{ text: "ok", algn: "ctr" },
				],
			},
		],
	}),
]);

const slide5 = slideXml([
	shape({
		name: "Title",
		ph: { type: "title" },
		text: { anchor: "b", paragraphs: [{ runs: [{ t: "Images and cropping" }] }] },
	}),
	picture({ name: "Full image", x: 120, y: 200, w: 440, h: 275, embed: "rId2" }),
	picture({ name: "Cropped image", x: 620, y: 200, w: 440, h: 275, embed: "rId2", crop: { l: 25, t: 20, r: 25, b: 20 } }),
	shape({
		name: "Caption",
		x: 120,
		y: 500,
		w: 940,
		h: 40,
		fill: "none",
		text: {
			paragraphs: [
				{ bullet: "none", runs: [{ t: "Left: the whole image. Right: the same image with a:srcRect insets applied.", sz: 15, color: "6B7B8C", i: true }] },
			],
		},
	}),
]);

const slide6 = slideXml([
	shape({
		name: "Title",
		ph: { type: "title" },
		text: { anchor: "b", paragraphs: [{ runs: [{ t: "Text alignment, breaks and links" }] }] },
	}),
	shape({
		name: "Left box",
		x: 80,
		y: 190,
		w: 520,
		h: 220,
		fill: { solid: "F5F8FC" },
		line: { w: 1, fill: { solid: "D6DEE8" } },
		text: {
			anchor: "ctr",
			paragraphs: [
				{ algn: "l", bullet: "none", runs: [{ t: "Left aligned" }] },
				{ algn: "ctr", bullet: "none", runs: [{ t: "Centred" }] },
				{ algn: "r", bullet: "none", runs: [{ t: "Right aligned" }] },
			],
		},
	}),
	shape({
		name: "Right box",
		x: 640,
		y: 190,
		w: 560,
		h: 220,
		fill: "none",
		text: {
			anchor: "ctr",
			paragraphs: [
				{
					bullet: "none",
					runs: [
						{ t: "A line break inside one paragraph:" },
						{ br: true },
						{ t: "…continues here without a new bullet." },
					],
				},
				{ bullet: "none", runs: [{ t: "" }] },
				{
					bullet: "none",
					runs: [
						{ t: "And a hyperlink to " },
						{ t: "obsidian.md", link: "rId2", color: "1971C2", u: true },
					],
				},
			],
		},
	}),
	shape({
		name: "Vertical anchors",
		x: 80,
		y: 440,
		w: 1120,
		h: 120,
		fill: { scheme: "accent1", lumMod: 20, lumOff: 80 },
		geom: "roundRect",
		text: {
			anchor: "ctr",
			paragraphs: [
				{ algn: "ctr", bullet: "none", runs: [{ t: "Vertically centred in a rounded rectangle", sz: 20, b: true, color: "1A3F7A" }] },
				{ algn: "ctr", bullet: "none", runs: [{ t: "This slide has speaker notes — press N to see them.", sz: 15, color: "3D5A80" }] },
			],
		},
	}),
]);

// ------------------------------------------------------------------ package

const image = encodePng(480, 300, (x, y) => {
	// A soft diagonal gradient with a grid, so cropping is obvious at a glance.
	const t = (x / 480 + y / 300) / 2;
	const grid = x % 60 < 2 || y % 60 < 2 ? -28 : 0;
	return [
		Math.max(0, Math.min(255, Math.round(40 + t * 120 + grid))),
		Math.max(0, Math.min(255, Math.round(90 + t * 110 + grid))),
		Math.max(0, Math.min(255, Math.round(200 - t * 60 + grid))),
	];
});

const REL = "officeDocument/2006/relationships";
const slides = [slide1, slide2, slide3, slide4, slide5, slide6];

const files = {
	"[Content_Types].xml":
		`${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
		'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
		'<Default Extension="xml" ContentType="application/xml"/>' +
		'<Default Extension="png" ContentType="image/png"/>' +
		'<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
		'<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
		'<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
		'<Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
		slides
			.map(
				(_, i) =>
					`<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
			)
			.join("") +
		'<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>' +
		'<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>' +
		'<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
		'<Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
		'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
		'<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
		"</Types>",

	"_rels/.rels": rels([
		{ id: "rId1", type: `${REL}/officeDocument`, target: "ppt/presentation.xml" },
		{ id: "rId2", type: "package/2006/relationships/metadata/core-properties", target: "docProps/core.xml" },
		{ id: "rId3", type: `${REL}/extended-properties`, target: "docProps/app.xml" },
	]),

	"docProps/core.xml":
		`${DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
		'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
		'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
		"<dc:title>PPTX Viewer fixture</dc:title><dc:creator>obsidian-pptx-viewer</dc:creator>" +
		"<cp:lastModifiedBy>obsidian-pptx-viewer</cp:lastModifiedBy></cp:coreProperties>",

	"docProps/app.xml":
		`${DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ` +
		'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
		`<Application>obsidian-pptx-viewer</Application><Slides>${slides.length}</Slides></Properties>`,

	"ppt/presentation.xml":
		`${DECL}<p:presentation ${NS} saveSubsetFonts="1">` +
		'<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
		'<p:notesMasterIdLst><p:notesMasterId r:id="rId100"/></p:notesMasterIdLst>' +
		"<p:sldIdLst>" +
		slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${10 + i}"/>`).join("") +
		"</p:sldIdLst>" +
		`<p:sldSz cx="${px(SLIDE_W)}" cy="${px(SLIDE_H)}"/><p:notesSz cx="${px(SLIDE_H)}" cy="${px(SLIDE_W)}"/>` +
		'<p:defaultTextStyle><a:lvl1pPr><a:defRPr sz="1800"><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle>' +
		"</p:presentation>",

	"ppt/_rels/presentation.xml.rels": rels([
		{ id: "rId1", type: `${REL}/slideMaster`, target: "slideMasters/slideMaster1.xml" },
		{ id: "rId100", type: `${REL}/notesMaster`, target: "notesMasters/notesMaster1.xml" },
		...slides.map((_, i) => ({
			id: `rId${10 + i}`,
			type: `${REL}/slide`,
			target: `slides/slide${i + 1}.xml`,
		})),
		{ id: "rId200", type: `${REL}/theme`, target: "theme/theme1.xml" },
	]),

	"ppt/slideMasters/slideMaster1.xml": MASTER,
	"ppt/slideMasters/_rels/slideMaster1.xml.rels": rels([
		{ id: "rId1", type: `${REL}/slideLayout`, target: "../slideLayouts/slideLayout1.xml" },
		{ id: "rId2", type: `${REL}/slideLayout`, target: "../slideLayouts/slideLayout2.xml" },
		{ id: "rId3", type: `${REL}/theme`, target: "../theme/theme1.xml" },
	]),

	"ppt/slideLayouts/slideLayout1.xml": TITLE_LAYOUT,
	"ppt/slideLayouts/_rels/slideLayout1.xml.rels": rels([
		{ id: "rId1", type: `${REL}/slideMaster`, target: "../slideMasters/slideMaster1.xml" },
	]),
	"ppt/slideLayouts/slideLayout2.xml": CONTENT_LAYOUT,
	"ppt/slideLayouts/_rels/slideLayout2.xml.rels": rels([
		{ id: "rId1", type: `${REL}/slideMaster`, target: "../slideMasters/slideMaster1.xml" },
	]),

	"ppt/notesMasters/notesMaster1.xml": NOTES_MASTER,
	"ppt/notesMasters/_rels/notesMaster1.xml.rels": rels([
		{ id: "rId1", type: `${REL}/theme`, target: "../theme/theme2.xml" },
	]),

	"ppt/notesSlides/notesSlide1.xml": notesSlideXml([
		"These notes come from ppt/notesSlides/notesSlide1.xml.",
		"They prove the notes pane and the Markdown extraction both pick them up.",
	]),
	"ppt/notesSlides/_rels/notesSlide1.xml.rels": rels([
		{ id: "rId1", type: `${REL}/notesMaster`, target: "../notesMasters/notesMaster1.xml" },
		{ id: "rId2", type: `${REL}/slide`, target: "../slides/slide6.xml" },
	]),

	"ppt/theme/theme1.xml": THEME,
	"ppt/theme/theme2.xml": THEME,
	"ppt/media/image1.png": new Uint8Array(image),
};

slides.forEach((xml, i) => {
	const n = i + 1;
	files[`ppt/slides/slide${n}.xml`] = xml;
	const layout = n === 1 ? "slideLayout1.xml" : "slideLayout2.xml";
	const entries = [{ id: "rId1", type: `${REL}/slideLayout`, target: `../slideLayouts/${layout}` }];
	if (n === 5) entries.push({ id: "rId2", type: `${REL}/image`, target: "../media/image1.png" });
	if (n === 6) {
		entries.push({ id: "rId2", type: `${REL}/hyperlink`, target: "https://obsidian.md", external: true });
		entries.push({ id: "rId3", type: `${REL}/notesSlide`, target: "../notesSlides/notesSlide1.xml" });
	}
	files[`ppt/slides/_rels/slide${n}.xml.rels`] = rels(entries);
});

const encoder = new TextEncoder();
const zipInput = {};
for (const [path, content] of Object.entries(files)) {
	zipInput[path] = typeof content === "string" ? encoder.encode(content) : content;
}

const outfile = process.argv[2] ?? "tests/fixtures/sample.pptx";
mkdirSync(dirname(outfile), { recursive: true });
writeFileSync(outfile, zipSync(zipInput, { level: 6 }));
console.log(`Wrote ${outfile} (${slides.length} slides)`);
