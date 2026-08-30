/**
 * A browser-shaped global environment for running the renderer under Node.
 *
 * The parser is pure data work, but the renderer and the edit write-back both
 * operate on real DOM, so testing them honestly means giving them a real DOM
 * rather than a hand-rolled stand-in that could quietly disagree with Chromium.
 */
import { JSDOM } from "jsdom";

export function installDom(): void {
	const dom = new JSDOM("<!doctype html><html><body></body></html>");
	const win = dom.window as unknown as Record<string, unknown>;
	const g = globalThis as unknown as Record<string, unknown>;

	for (const key of [
		"window",
		"document",
		"DOMParser",
		"XMLSerializer",
		"Node",
		"Element",
		"HTMLElement",
		"HTMLImageElement",
		"HTMLAnchorElement",
		"SVGElement",
		"Image",
		"getComputedStyle",
	]) {
		g[key] = key === "window" ? dom.window : win[key];
	}

	// Obsidian adds these to HTMLElement; the renderer uses them freely.
	const proto = (dom.window as unknown as { HTMLElement: { prototype: HTMLElement } }).HTMLElement
		.prototype as HTMLElement & Record<string, unknown>;
	proto.addClass = function (this: HTMLElement, ...classes: string[]) {
		this.classList.add(...classes);
		return this;
	};
	proto.removeClass = function (this: HTMLElement, ...classes: string[]) {
		this.classList.remove(...classes);
		return this;
	};
	proto.hasClass = function (this: HTMLElement, cls: string) {
		return this.classList.contains(cls);
	};
	proto.toggleClass = function (this: HTMLElement, cls: string, value: boolean) {
		this.classList.toggle(cls, value);
		return this;
	};
	proto.setText = function (this: HTMLElement, text: string) {
		this.textContent = text;
	};
	proto.empty = function (this: HTMLElement) {
		while (this.firstChild) this.removeChild(this.firstChild);
	};
	proto.detach = function (this: HTMLElement) {
		this.parentNode?.removeChild(this);
	};

	// Object URLs have no meaning outside a browser; the renderer only needs a string.
	const url = g.URL as unknown as Record<string, unknown>;
	url.createObjectURL = () => "blob:test";
	url.revokeObjectURL = () => undefined;
}
