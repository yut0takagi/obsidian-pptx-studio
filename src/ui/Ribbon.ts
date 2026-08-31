import { Menu, setIcon, setTooltip } from "obsidian";
import { t } from "../i18n";

export interface RibbonButton {
	kind: "button";
	icon?: string;
	label?: string;
	tooltip: string;
	onClick: () => void;
	isEnabled?: () => boolean;
	isActive?: () => boolean;
}

export interface RibbonMenu {
	kind: "menu";
	icon?: string;
	label?: string;
	tooltip: string;
	build: (menu: Menu) => void;
	isEnabled?: () => boolean;
}

export interface RibbonSelect {
	kind: "select";
	tooltip: string;
	width?: string;
	options: () => { value: string; label: string }[];
	value: () => string;
	onChange: (value: string) => void;
	isEnabled?: () => boolean;
}

export interface RibbonColor {
	kind: "color";
	tooltip: string;
	icon: string;
	value: () => string | null;
	onChange: (value: string | null) => void;
	/** Offer a "no fill" / "no outline" entry. */
	allowNone?: boolean;
	isEnabled?: () => boolean;
}

export interface RibbonNumber {
	kind: "number";
	tooltip: string;
	label?: string;
	step?: number;
	width?: string;
	/** null when the selection has no single value. */
	value: () => number | null;
	onChange: (value: number) => void;
	isEnabled?: () => boolean;
}

export interface RibbonSeparator {
	kind: "separator";
}

export type RibbonItem =
	| RibbonButton
	| RibbonMenu
	| RibbonSelect
	| RibbonColor
	| RibbonNumber
	| RibbonSeparator;

export interface RibbonGroup {
	title: string;
	items: RibbonItem[];
}

export interface RibbonTab {
	id: string;
	title: string;
	groups: RibbonGroup[];
	/** Contextual tabs (Table) only appear while they apply to the selection. */
	visible?: () => boolean;
}

/**
 * A tabbed ribbon in the PowerPoint idiom.
 *
 * Controls are declared as data and re-evaluated on every `update()`, so a
 * command that becomes unavailable — align with nothing selected, ungroup on a
 * plain shape — greys itself out without any per-control bookkeeping.
 */
export interface RibbonOptions {
	/** Start collapsed, and report the state back so it can be remembered. */
	collapsed?: boolean;
	initialTab?: string;
	onStateChange?: (state: { collapsed: boolean; tab: string }) => void;
}

export class Ribbon {
	private readonly root: HTMLElement;
	private readonly tabsEl: HTMLElement;
	private readonly bodyEl: HTMLElement;
	private readonly refreshers: (() => void)[] = [];
	private activeTab: string;
	private collapsed: boolean;
	/** The tab to return to once a contextual tab stops applying. */
	private tabBeforeContext: string | null = null;
	private visibleContextual = new Set<string>();
	private resizeObserver: ResizeObserver | null = null;

	constructor(
		containerEl: HTMLElement,
		private readonly tabs: RibbonTab[],
		private readonly options: RibbonOptions = {},
	) {
		this.root = containerEl.createDiv({ cls: "pptx-ribbon" });
		this.tabsEl = this.root.createDiv({ cls: "pptx-ribbon-tabs" });
		this.bodyEl = this.root.createDiv({ cls: "pptx-ribbon-body" });
		this.collapsed = options.collapsed ?? false;
		const wanted = options.initialTab;
		this.activeTab = tabs.some((tab) => tab.id === wanted && !tab.visible)
			? (wanted as string)
			: (tabs[0]?.id ?? "");
		this.buildTabs();
		this.showTab(this.activeTab);
		this.applyCollapsed();

		// In a narrow pane the ribbon drops its labels rather than scrolling, which
		// is the difference between "cramped" and "unusable".
		this.resizeObserver = new ResizeObserver(() => {
			this.root.toggleClass("is-compact", this.root.clientWidth < 760);
		});
		this.resizeObserver.observe(this.root);
	}

	private buildTabs(): void {
		for (const tab of this.tabs) {
			const el = this.tabsEl.createEl("button", { cls: "pptx-ribbon-tab", text: tab.title });
			el.dataset.tab = tab.id;
			if (tab.visible) el.addClass("is-contextual");
			el.addEventListener("mousedown", (event) => event.preventDefault());
			el.addEventListener("click", () => {
				// Choosing a tab by hand cancels the automatic return.
				this.tabBeforeContext = null;
				if (this.collapsed) this.setCollapsed(false);
				this.showTab(tab.id);
			});
		}
		const toggle = this.tabsEl.createEl("button", { cls: "pptx-ribbon-collapse clickable-icon" });
		toggle.addEventListener("mousedown", (event) => event.preventDefault());
		toggle.addEventListener("click", () => this.setCollapsed(!this.collapsed));
		this.collapseButton = toggle;
	}

	private collapseButton: HTMLElement | null = null;

	setCollapsed(collapsed: boolean): void {
		this.collapsed = collapsed;
		this.applyCollapsed();
		this.options.onStateChange?.({ collapsed, tab: this.activeTab });
	}

	private applyCollapsed(): void {
		this.root.toggleClass("is-collapsed", this.collapsed);
		if (this.collapseButton) {
			setIcon(this.collapseButton, this.collapsed ? "chevron-down" : "chevron-up");
			setTooltip(this.collapseButton, this.collapsed ? t("ribbon.expand") : t("ribbon.collapse"));
		}
	}

	showTab(id: string): void {
		this.activeTab = id;
		for (const el of Array.from(this.tabsEl.children)) {
			const tabEl = el as HTMLElement;
			tabEl.toggleClass("is-active", tabEl.dataset.tab === id);
		}
		this.bodyEl.empty();
		this.refreshers.length = 0;
		const tab = this.tabs.find((entry) => entry.id === id);
		if (!tab) return;
		// Contextual tabs are not worth remembering: they only apply while
		// something specific is selected, which will not be true next time.
		if (!tab.visible) this.options.onStateChange?.({ collapsed: this.collapsed, tab: id });
		tab.groups.forEach((group, index) => {
			if (index > 0) this.bodyEl.createDiv({ cls: "pptx-ribbon-divider" });
			this.buildGroup(group);
		});
		this.update();
	}

	get currentTab(): string {
		return this.activeTab;
	}

	private buildGroup(group: RibbonGroup): void {
		const el = this.bodyEl.createDiv({ cls: "pptx-ribbon-group" });
		const row = el.createDiv({ cls: "pptx-ribbon-row" });
		for (const item of group.items) this.buildItem(row, item);
		el.createDiv({ cls: "pptx-ribbon-group-title", text: group.title });
	}

	private buildItem(parent: HTMLElement, item: RibbonItem): void {
		switch (item.kind) {
			case "separator":
				parent.createDiv({ cls: "pptx-ribbon-sep" });
				return;

			case "button": {
				const el = parent.createEl("button", { cls: "pptx-ribbon-btn" });
				// Keeping focus on the canvas is what lets the arrow keys keep working
				// after a click up here; a focused button would swallow them.
				el.addEventListener("mousedown", (event) => event.preventDefault());
				if (item.icon) setIcon(el.createSpan({ cls: "pptx-ribbon-icon" }), item.icon);
				if (item.label) el.createSpan({ cls: "pptx-ribbon-label", text: item.label });
				setTooltip(el, item.tooltip);
				el.addEventListener("click", (event) => {
					event.preventDefault();
					if (el.hasClass("is-disabled")) return;
					item.onClick();
				});
				this.refreshers.push(() => {
					el.toggleClass("is-disabled", item.isEnabled ? !item.isEnabled() : false);
					el.toggleClass("is-active", item.isActive ? item.isActive() : false);
				});
				return;
			}

			case "menu": {
				const el = parent.createEl("button", { cls: "pptx-ribbon-btn has-menu" });
				el.addEventListener("mousedown", (event) => event.preventDefault());
				if (item.icon) setIcon(el.createSpan({ cls: "pptx-ribbon-icon" }), item.icon);
				if (item.label) el.createSpan({ cls: "pptx-ribbon-label", text: item.label });
				setIcon(el.createSpan({ cls: "pptx-ribbon-caret" }), "chevron-down");
				setTooltip(el, item.tooltip);
				el.addEventListener("click", (event) => {
					event.preventDefault();
					if (el.hasClass("is-disabled")) return;
					const menu = new Menu();
					item.build(menu);
					const rect = el.getBoundingClientRect();
					menu.showAtPosition({ x: rect.left, y: rect.bottom });
				});
				this.refreshers.push(() => {
					el.toggleClass("is-disabled", item.isEnabled ? !item.isEnabled() : false);
				});
				return;
			}

			case "select": {
				const el = parent.createEl("select", { cls: "pptx-ribbon-select dropdown" });
				if (item.width) el.style.width = item.width;
				setTooltip(el, item.tooltip);
				let signature = "";
				el.addEventListener("change", () => item.onChange(el.value));
				this.refreshers.push(() => {
					const options = item.options();
					const next = options.map((o) => o.value).join(" ");
					if (next !== signature) {
						signature = next;
						el.empty();
						for (const option of options) {
							el.createEl("option", { value: option.value, text: option.label });
						}
					}
					// A mixed selection has no single value; leave the control blank
					// rather than claiming the first shape's value applies to all.
					const value = item.value();
					el.value = value;
					if (el.value !== value) el.selectedIndex = -1;
					el.disabled = item.isEnabled ? !item.isEnabled() : false;
				});
				return;
			}

			case "number": {
				const wrap = parent.createDiv({ cls: "pptx-ribbon-number" });
				if (item.label) wrap.createSpan({ cls: "pptx-ribbon-number-label", text: item.label });
				const el = wrap.createEl("input", { type: "number" });
				el.step = String(item.step ?? 1);
				if (item.width) el.style.width = item.width;
				setTooltip(el, item.tooltip);
				const commit = (): void => {
					const value = Number(el.value);
					if (Number.isFinite(value)) item.onChange(value);
				};
				el.addEventListener("change", commit);
				el.addEventListener("keydown", (event) => {
					event.stopPropagation();
					if (event.key === "Enter") commit();
				});
				this.refreshers.push(() => {
					// Never fight the user's typing: only write while unfocused.
					if (document.activeElement !== el) {
						const value = item.value();
						el.value = value === null ? "" : String(Math.round(value * 10) / 10);
					}
					el.disabled = item.isEnabled ? !item.isEnabled() : false;
				});
				return;
			}

			case "color": {
				const el = parent.createEl("button", { cls: "pptx-ribbon-btn pptx-ribbon-color" });
				el.addEventListener("mousedown", (event) => event.preventDefault());
				setIcon(el.createSpan({ cls: "pptx-ribbon-icon" }), item.icon);
				const swatch = el.createDiv({ cls: "pptx-ribbon-swatch" });
				setTooltip(el, item.tooltip);
				el.addEventListener("click", (event) => {
					event.preventDefault();
					if (el.hasClass("is-disabled")) return;
					openColorPopup(el, item);
				});
				this.refreshers.push(() => {
					const value = item.value();
					swatch.setCssStyles({ background: value ?? "transparent" });
					swatch.toggleClass("is-none", value === null);
					el.toggleClass("is-disabled", item.isEnabled ? !item.isEnabled() : false);
				});
				return;
			}
		}
	}

	/** Re-evaluate every control's enabled state and displayed value. */
	update(): void {
		let activeStillVisible = true;
		const nowVisible = new Set<string>();

		for (const el of Array.from(this.tabsEl.children)) {
			const tabEl = el as HTMLElement;
			const id = tabEl.dataset.tab;
			if (!id) continue;
			const tab = this.tabs.find((entry) => entry.id === id);
			const contextual = Boolean(tab?.visible);
			const visible = tab?.visible ? tab.visible() : true;
			tabEl.toggleClass("is-hidden", !visible);
			if (contextual && visible) nowVisible.add(id);
			if (!visible && id === this.activeTab) activeStillVisible = false;
		}

		// A contextual tab appearing means the selection just became something it
		// applies to, so switch to it — and switch back when it goes away.
		for (const id of nowVisible) {
			if (!this.visibleContextual.has(id) && this.activeTab !== id) {
				this.tabBeforeContext = this.activeTab;
				this.showTab(id);
				break;
			}
		}
		if (!activeStillVisible || (this.tabBeforeContext && !nowVisible.has(this.activeTab))) {
			const fallback =
				this.tabBeforeContext ?? this.tabs.find((entry) => !entry.visible)?.id ?? null;
			this.tabBeforeContext = null;
			if (fallback && fallback !== this.activeTab) {
				this.showTab(fallback);
				this.visibleContextual = nowVisible;
				return;
			}
		}
		this.visibleContextual = nowVisible;

		for (const refresh of this.refreshers) refresh();
	}

	destroy(): void {
		this.resizeObserver?.disconnect();
		this.root.detach();
	}
}

/** Theme-neutral palette offered by every colour control. */
const SWATCHES = [
	"#000000",
	"#404040",
	"#808080",
	"#bfbfbf",
	"#ffffff",
	"#c92a2a",
	"#e8590c",
	"#f59f00",
	"#2f9e44",
	"#12b886",
	"#1971c2",
	"#2f6fed",
	"#7048e8",
	"#c2255c",
	"#e64980",
];

/** The colour picker, shared with the floating selection toolbar. */
export function openColorPopup(anchor: HTMLElement, item: RibbonColor): void {
	const rect = anchor.getBoundingClientRect();
	const popup = document.body.createDiv({ cls: "pptx-color-popup" });
	Object.assign(popup.style, {
		position: "fixed",
		left: `${rect.left}px`,
		top: `${rect.bottom + 4}px`,
	});

	const close = (): void => {
		document.removeEventListener("mousedown", onOutside, true);
		popup.remove();
	};
	const onOutside = (event: MouseEvent): void => {
		if (!popup.contains(event.target as Node)) close();
	};

	const grid = popup.createDiv({ cls: "pptx-color-grid" });
	for (const color of SWATCHES) {
		const cell = grid.createDiv({ cls: "pptx-color-cell" });
		cell.style.background = color;
		setTooltip(cell, color);
		cell.addEventListener("click", () => {
			item.onChange(color);
			close();
		});
	}

	if (item.allowNone) {
		const none = popup.createDiv({ cls: "pptx-color-option" });
		none.setText(t("common.none"));
		none.addEventListener("click", () => {
			item.onChange(null);
			close();
		});
	}

	const custom = popup.createDiv({ cls: "pptx-color-option" });
	const input = custom.createEl("input", { type: "color" });
	input.value = item.value() ?? "#000000";
	custom.createSpan({ text: t("common.custom") });
	input.addEventListener("change", () => {
		item.onChange(input.value);
		close();
	});

	window.setTimeout(() => document.addEventListener("mousedown", onOutside, true), 0);
}
