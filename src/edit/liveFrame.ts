/**
 * Writing a frame straight onto the rendered element, mid-gesture.
 *
 * A drag cannot afford to re-derive the slide on every pointer move, so the
 * element is moved directly and the model catches up when the gesture ends.
 */
import type { Frame, Shape } from "../pptx/types";

export function place(el: HTMLElement, frame: Frame): void {
	el.setCssStyles({
		left: `${frame.x}px`,
		top: `${frame.y}px`,
		width: `${frame.w}px`,
		height: `${frame.h}px`,
	});
}

/** Move a shape's element, taking a group's children with it. */
export function applyLive(el: HTMLElement, shape: Shape, frame: Frame): void {
	place(el, frame);
	if (shape.kind === "group") {
		const inner = el.firstElementChild;
		if (inner?.instanceOf(HTMLElement)) {
			const sx = shape.childOffset.w > 0 ? frame.w / shape.childOffset.w : 1;
			const sy = shape.childOffset.h > 0 ? frame.h / shape.childOffset.h : 1;
			inner.setCssStyles({
				transform: `scale(${sx}, ${sy}) translate(${-shape.childOffset.x}px, ${-shape.childOffset.y}px)`,
			});
		}
	}
}
