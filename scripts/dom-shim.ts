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

	g.createEl = (tag: string) => document.createElement(tag);
	g.createSvg = (tag: string) => document.createElementNS("http://www.w3.org/2000/svg", tag);
	g.createDiv = () => document.createElement("div");
	g.createSpan = () => document.createElement("span");

	Object.defineProperty(dom.window.Object.prototype, "instanceOf", {
		configurable: true,
		value: function (this: unknown, type: unknown) {
			return this instanceof (type as new (...args: never[]) => unknown);
		},
	});

	// Obsidian adds these to elements; the renderer uses them freely.
	const proto = (dom.window as unknown as { Element: { prototype: Element } }).Element
		.prototype as Element & Record<string, unknown>;
	proto.addClass = function (this: Element, ...classes: string[]) {
		this.classList.add(...classes);
		return this;
	};
	proto.removeClass = function (this: Element, ...classes: string[]) {
		this.classList.remove(...classes);
		return this;
	};
	proto.hasClass = function (this: Element, cls: string) {
		return this.classList.contains(cls);
	};
	proto.toggleClass = function (this: Element, cls: string, value: boolean) {
		this.classList.toggle(cls, value);
		return this;
	};
	proto.setText = function (this: Element, text: string) {
		this.textContent = text;
	};
	proto.empty = function (this: Element) {
		while (this.firstChild) this.removeChild(this.firstChild);
	};
	proto.detach = function (this: Element) {
		this.parentNode?.removeChild(this);
	};
	proto.setCssStyles = function (this: HTMLElement | SVGElement, styles: Record<string, string>) {
		Object.assign(this.style, styles);
	};

	const htmlProto = (dom.window as unknown as { HTMLElement: { prototype: HTMLElement } })
		.HTMLElement.prototype as HTMLElement & Record<string, unknown>;
	htmlProto.createEl = function (this: HTMLElement, tag: string, options?: { cls?: string; text?: string }) {
		const child = document.createElement(tag);
		if (options?.cls) child.addClass(options.cls);
		if (options?.text !== undefined) child.setText(options.text);
		this.appendChild(child);
		return child;
	};
	htmlProto.createDiv = function (this: HTMLElement, options?: { cls?: string; text?: string }) {
		return this.createEl("div", options);
	};
	htmlProto.createSpan = function (this: HTMLElement, options?: { cls?: string; text?: string }) {
		return this.createEl("span", options);
	};

	// Object URLs have no meaning outside a browser; the renderer only needs a string.
	const url = g.URL as unknown as Record<string, unknown>;
	url.createObjectURL = () => "blob:test";
	url.revokeObjectURL = () => undefined;
}
