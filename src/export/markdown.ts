import type { Deck, Paragraph, Shape, Slide, TextBody } from "../pptx/types";

export interface MarkdownOptions {
	includeNotes: boolean;
	includeSourceLink: boolean;
	/** Vault path of the source deck, used for the back-link. */
	sourcePath: string;
}

/**
 * Flatten a deck into a Markdown outline. The point is searchability: every word
 * on every slide ends up in a note that Obsidian indexes, links and backlinks.
 */
export function deckToMarkdown(deck: Deck, options: MarkdownOptions): string {
	const out: string[] = [];
	out.push(`# ${deck.title}`, "");
	if (options.includeSourceLink && options.sourcePath) {
		out.push(`Source: [[${options.sourcePath}]]`, "");
	}

	for (const slide of deck.slides) {
		out.push(...slideToMarkdown(slide, options));
	}

	return out.join("\n").replace(/\n{3,}/g, "\n\n").trimEnd() + "\n";
}

function slideToMarkdown(slide: Slide, options: MarkdownOptions): string[] {
	// Only the slide's own shapes; layout and master artwork is template chrome.
	const own = slide.shapes.slice(slide.templateShapes);
	const { text: title, shape: titleShape } = findTitle(own);
	const out: string[] = [];
	out.push(`## Slide ${slide.index}${title ? ` — ${title}` : ""}`, "");

	for (const shape of own) {
		if (shape === titleShape) continue;
		out.push(...shapeToMarkdown(shape, title));
	}

	if (options.includeNotes && slide.notes) {
		out.push("> [!note] Speaker notes");
		for (const line of slide.notes.split("\n")) {
			out.push(`> ${line}`);
		}
		out.push("");
	}

	out.push("");
	return out;
}

/** The slide's title placeholder, so it becomes the heading rather than a bullet. */
function findTitle(shapes: Shape[]): { text: string; shape: Shape | null } {
	for (const shape of shapes) {
		if (shape.kind !== "shape" || !shape.text) continue;
		if (shape.placeholder === "title" || shape.placeholder === "ctrTitle") {
			const text = bodyToLines(shape.text).join(" ").trim();
			if (text) return { text, shape };
		}
	}
	return { text: "", shape: null };
}

function shapeToMarkdown(shape: Shape, title: string): string[] {
	switch (shape.kind) {
		case "shape": {
			if (!shape.text) return [];
			const lines: string[] = [];
			for (const para of shape.text.paragraphs) {
				const text = paragraphText(para);
				if (!text) continue;
				if (text === title) continue;
				lines.push(para.bullet ? `${"  ".repeat(para.level)}- ${text}` : text);
			}
			return lines.length ? [...lines, ""] : [];
		}
		case "table":
			return tableToMarkdown(shape.table);
		case "chart": {
			const parts = [`**${shape.title || `${shape.chartType} chart`}**`];
			if (shape.series.length) parts.push(`Series: ${shape.series.join(", ")}`);
			if (shape.categories.length) parts.push(`Categories: ${shape.categories.join(", ")}`);
			return [...parts, ""];
		}
		case "group":
			return shape.children.flatMap((kid) => shapeToMarkdown(kid, title));
		case "image":
			return shape.label && shape.label !== "Image" ? [`![${shape.label}]()`, ""] : [];
		case "line":
			return [];
	}
}

function tableToMarkdown(table: import("../pptx/types").Table): string[] {
	if (table.rows.length === 0) return [];
	const rows = table.rows.map((row) =>
		row.cells.map((cell) => bodyToLines(cell.text).join(" ").replace(/\|/g, "\\|").trim()),
	);
	const width = Math.max(...rows.map((r) => r.length));
	const pad = (row: string[]) => {
		const copy = [...row];
		while (copy.length < width) copy.push("");
		return copy;
	};
	const out = [`| ${pad(rows[0]).join(" | ")} |`, `| ${Array(width).fill("---").join(" | ")} |`];
	for (const row of rows.slice(1)) {
		out.push(`| ${pad(row).join(" | ")} |`);
	}
	out.push("");
	return out;
}

function paragraphText(para: Paragraph): string {
	return para.runs
		.map((r) => r.text)
		.join("")
		.replace(/\n/g, " ")
		.trim();
}

function bodyToLines(body: TextBody | null): string[] {
	if (!body) return [];
	return body.paragraphs.map(paragraphText).filter((t) => t !== "");
}

/** Plain text of a whole slide, used for the view's search box. */
export function slideSearchText(slide: Slide): string {
	const own = slide.shapes.slice(slide.templateShapes);
	const collect = (shape: Shape): string[] => {
		switch (shape.kind) {
			case "shape":
				return bodyToLines(shape.text);
			case "table":
				return shape.table.rows.flatMap((r) => r.cells.flatMap((c) => bodyToLines(c.text)));
			case "chart":
				return [shape.title, ...shape.series, ...shape.categories];
			case "group":
				return shape.children.flatMap(collect);
			default:
				return [];
		}
	};
	return [...own.flatMap(collect), slide.notes].join(" ").toLowerCase();
}
