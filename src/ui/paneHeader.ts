import { setIcon, setTooltip } from "obsidian";
import { t } from "../i18n";

export interface PaneHeaderOptions {
	collapsed: boolean;
	onToggle: (collapsed: boolean) => void;
}

/**
 * The title bar of a side-pane section, which folds the section away.
 *
 * The whole bar is the target rather than just the chevron: the sections are
 * small, and a title you have to aim past to hit a 14-pixel arrow is a title
 * that gets missed. Returns the bar, so a section can add its own controls.
 */
export function buildPaneHeader(
	root: HTMLElement,
	title: string,
	options: PaneHeaderOptions,
): HTMLElement {
	let collapsed = options.collapsed;
	const head = root.createDiv({ cls: "pptx-pane-title" });
	const chevron = head.createSpan({ cls: "pptx-pane-chevron" });
	head.createSpan({ cls: "pptx-pane-heading", text: title });

	const apply = (): void => {
		root.toggleClass("is-collapsed", collapsed);
		setIcon(chevron, collapsed ? "chevron-right" : "chevron-down");
		setTooltip(head, collapsed ? t("pane.expand") : t("pane.collapse"));
	};
	apply();

	head.addEventListener("click", () => {
		collapsed = !collapsed;
		apply();
		options.onToggle(collapsed);
	});
	return head;
}
