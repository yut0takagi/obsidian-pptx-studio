export function getLanguage(): string {
	return "en";
}

/** Icons are SVG in Obsidian; the tests only need to see which one was asked for. */
export function setIcon(el: HTMLElement, icon: string): void {
	el.dataset.icon = icon;
}
