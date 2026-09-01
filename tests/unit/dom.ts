/**
 * A DOM parser for the tests, and nothing else.
 *
 * The parsers under test take `Element`s, so they need something that can make
 * one — but not the Obsidian element helpers `scripts/dom-shim.ts` installs for
 * the renderer. Keeping this separate means a unit test cannot accidentally
 * lean on behaviour that only exists because the renderer needed it.
 */
import { JSDOM } from "jsdom";

export function installDomParser(): void {
	const dom = new JSDOM("<!doctype html><html><body></body></html>");
	const win = dom.window as unknown as Record<string, unknown>;
	const g = globalThis as unknown as Record<string, unknown>;
	for (const key of ["DOMParser", "XMLSerializer", "Node", "Element", "Document"]) {
		g[key] = win[key];
	}
}
