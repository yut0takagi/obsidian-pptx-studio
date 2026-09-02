/**
 * The Animation tab: how a slide arrives, and how the shapes on it do.
 *
 * It lives apart from the other tabs because it is the only one whose state is
 * not in the model — transitions and the timing tree are read straight out of
 * the slide's XML each time the ribbon refreshes, since nothing else in the
 * plugin has reason to carry them.
 */
import { type StringKey, t } from "../i18n";
import {
	animateSelection,
	canReorderAnimations,
	clearSlideAnimations,
	moveSelectionAnimation,
	removeSelectionAnimation,
	selectedAnimations,
	setSlideTransition,
	slideAnimations,
	slideTransition,
} from "../edit/animationCommands";
import { selectedShapes } from "../edit/commands";
import { DIRECTED, type EffectKind, type TransitionKind, type Trigger } from "../ooxml/animation";
import type { RibbonItem, RibbonTab } from "./Ribbon";
import type { RibbonHost } from "./tabs";

export function buildAnimationTab(host: RibbonHost): RibbonTab {
	const hasSelection = () => (host.ctx() ? selectedShapes(host.ctx()!).length > 0 : false);

	// The trigger the next entrance will use. Changing it while a shape that is
	// already animated is selected re-applies that shape's effect, so the select
	// reads as the state of the selection rather than as a hidden preference.
	let pendingTrigger: Trigger = "click";

	const currentAnimation = () => selectedAnimations(host.ctx())[0] ?? null;
	const transitionOf = () => slideTransition(host.ctx());

	const TRANSITIONS: { value: string; key: StringKey }[] = [
		{ value: "none", key: "transition.none" },
		{ value: "fade", key: "transition.fade" },
		{ value: "cut", key: "transition.cut" },
		{ value: "push", key: "transition.push" },
		{ value: "wipe", key: "transition.wipe" },
		{ value: "cover", key: "transition.cover" },
		{ value: "dissolve", key: "transition.dissolve" },
		{ value: "split", key: "transition.split" },
		{ value: "randomBar", key: "transition.randomBar" },
	];

	const transitionItems: RibbonItem[] = [
		{
			kind: "select",
			tooltip: t("cmd.transition"),
			width: "10em",
			options: () => TRANSITIONS.map((entry) => ({ value: entry.value, label: t(entry.key) })),
			value: () => transitionOf()?.kind ?? "none",
			isEnabled: host.canEdit,
			onChange: (value) =>
				host.run((ctx) => {
					const current = transitionOf();
					setSlideTransition(
						ctx,
						value === "none"
							? null
							: {
									kind: value as TransitionKind,
									speed: current?.speed ?? "med",
									direction: current?.direction ?? null,
								},
						t("cmd.transition"),
					);
				}),
		},
		{
			kind: "select",
			tooltip: t("cmd.transitionSpeed"),
			width: "6em",
			options: () => [
				{ value: "slow", label: t("speed.slow") },
				{ value: "med", label: t("speed.med") },
				{ value: "fast", label: t("speed.fast") },
			],
			value: () => transitionOf()?.speed ?? "med",
			isEnabled: () => host.canEdit() && transitionOf() !== null,
			onChange: (value) =>
				host.run((ctx) => {
					const current = transitionOf();
					if (!current) return;
					setSlideTransition(
						ctx,
						{ ...current, speed: value as "slow" | "med" | "fast" },
						t("cmd.transitionSpeed"),
					);
				}),
		},
		{
			kind: "select",
			tooltip: t("cmd.transitionDirection"),
			width: "7em",
			options: () =>
				(DIRECTED[transitionOf()?.kind ?? ""] ?? []).map((dir) => ({
					value: dir,
					label: t(`dir.${dir}` as StringKey),
				})),
			value: () => transitionOf()?.direction ?? "",
			isEnabled: () => host.canEdit() && DIRECTED[transitionOf()?.kind ?? ""] !== undefined,
			onChange: (value) =>
				host.run((ctx) => {
					const current = transitionOf();
					if (!current) return;
					setSlideTransition(ctx, { ...current, direction: value }, t("cmd.transitionDirection"));
				}),
		},
	];

	const EFFECTS: { value: EffectKind; key: StringKey }[] = [
		{ value: "fade", key: "effect.fade" },
		{ value: "wipe", key: "effect.wipe" },
		{ value: "flyIn", key: "effect.flyIn" },
	];

	const animationItems: RibbonItem[] = [
		{
			kind: "menu",
			icon: "sparkles",
			label: t("cmd.animate"),
			tooltip: t("cmd.animate"),
			isEnabled: () => host.canEdit() && hasSelection(),
			build: (menu) => {
				for (const entry of EFFECTS) {
					menu.addItem((item) =>
						item
							.setTitle(t(entry.key))
							.setChecked(currentAnimation()?.effect === entry.value)
							.onClick(() =>
								host.run((ctx) =>
									animateSelection(ctx, entry.value, pendingTrigger, t("cmd.animate")),
								),
							),
					);
				}
			},
		},
		{
			kind: "select",
			tooltip: t("cmd.animateTrigger"),
			width: "10em",
			options: () => [
				{ value: "click", label: t("trigger.click") },
				{ value: "withPrev", label: t("trigger.withPrev") },
				{ value: "afterPrev", label: t("trigger.afterPrev") },
			],
			value: () => currentAnimation()?.trigger ?? pendingTrigger,
			isEnabled: () => host.canEdit() && hasSelection(),
			onChange: (value) => {
				pendingTrigger = value as Trigger;
				const existing = currentAnimation();
				if (!existing || existing.effect === "other") return;
				host.run((ctx) =>
					animateSelection(ctx, existing.effect as EffectKind, pendingTrigger, t("cmd.animateTrigger")),
				);
			},
		},
		{
			kind: "button",
			icon: "eraser",
			tooltip: t("cmd.animationRemove"),
			isEnabled: () => host.canEdit() && selectedAnimations(host.ctx()).length > 0,
			onClick: () => host.run((ctx) => removeSelectionAnimation(ctx, t("cmd.animationRemove"))),
		},
		{
			kind: "button",
			icon: "trash-2",
			tooltip: t("cmd.animationClear"),
			isEnabled: () => host.canEdit() && slideAnimations(host.ctx()).length > 0,
			onClick: () => host.run((ctx) => clearSlideAnimations(ctx, t("cmd.animationClear"))),
		},
	];

	// Reordering rebuilds the sequence, so it is offered only when every effect
	// on the slide is one this plugin knows how to write back.
	const canMove = (by: -1 | 1) => {
		if (!host.canEdit() || !canReorderAnimations(host.ctx())) return false;
		const mine = selectedAnimations(host.ctx());
		if (mine.length !== 1) return false;
		const at = mine[0].index;
		return by === -1 ? at > 0 : at < slideAnimations(host.ctx()).length - 1;
	};

	const animationOrderItems: RibbonItem[] = [
		{
			kind: "button",
			icon: "chevron-up",
			tooltip: t("cmd.animationEarlier"),
			isEnabled: () => canMove(-1),
			onClick: () => host.run((ctx) => moveSelectionAnimation(ctx, -1, t("cmd.animationEarlier"))),
		},
		{
			kind: "button",
			icon: "chevron-down",
			tooltip: t("cmd.animationLater"),
			isEnabled: () => canMove(1),
			onClick: () => host.run((ctx) => moveSelectionAnimation(ctx, 1, t("cmd.animationLater"))),
		},
	];

	return {
		id: "animation",
		title: t("tab.animation"),
		groups: [
			{ title: t("group.transition"), items: transitionItems },
			{ title: t("group.animation"), items: animationItems },
			{ title: t("group.animationOrder"), items: animationOrderItems },
		],
	};
}
