/**
 * Transitions, and the animation main sequence.
 *
 * A timing tree PowerPoint will not open is the failure that matters here, and
 * it is not one a reader spots: the tree is deep, the ids have to be unique,
 * and an empty `p:par` left behind by a deletion is as fatal as a missing one.
 * So these check the shape of what is written, not only what can be read back
 * out of it.
 */
import { strict as assert } from "node:assert";
import { beforeEach, describe, it } from "node:test";
import { installDomParser } from "./dom";

installDomParser();

import {
	addAnimation,
	canReorder,
	clearAnimations,
	moveAnimation,
	readAnimations,
	readTransition,
	removeAnimationAt,
	removeAnimationsFor,
	writeTransition,
} from "../../src/ooxml/animation";

const P = "http://schemas.openxmlformats.org/presentationml/2006/main";

let slide: Element;

beforeEach(() => {
	slide = new DOMParser().parseFromString(
		`<p:sld xmlns:p="${P}"><p:cSld><p:spTree/></p:cSld><p:clrMapOvr/></p:sld>`,
		"application/xml",
	).documentElement;
});

/** Local names of the slide's children, to check schema order. */
const childOrder = (): string[] => {
	const out: string[] = [];
	for (let n = slide.firstElementChild; n; n = n.nextElementSibling) out.push(n.localName);
	return out;
};

const all = (localName: string): Element[] =>
	Array.from(slide.getElementsByTagName("*")).filter((el) => el.localName === localName);

describe("transitions", () => {
	it("has none to start with", () => {
		assert.equal(readTransition(slide), null);
	});

	it("writes one and reads it back", () => {
		writeTransition(slide, { kind: "fade", speed: "fast", direction: null });
		assert.deepEqual(readTransition(slide), { kind: "fade", speed: "fast", direction: null });
	});

	it("gives a directional transition a direction, even when asked for none", () => {
		writeTransition(slide, { kind: "push", speed: "med", direction: null });
		assert.equal(readTransition(slide)?.direction, "l");
	});

	it("keeps a direction it was given", () => {
		writeTransition(slide, { kind: "wipe", speed: "med", direction: "d" });
		assert.equal(readTransition(slide)?.direction, "d");
	});

	it("refuses a direction the transition does not travel in", () => {
		writeTransition(slide, { kind: "push", speed: "med", direction: "sideways" });
		assert.equal(readTransition(slide)?.direction, "l");
	});

	it("replaces rather than accumulating", () => {
		writeTransition(slide, { kind: "fade", speed: "med", direction: null });
		writeTransition(slide, { kind: "cut", speed: "med", direction: null });
		assert.equal(all("transition").length, 1);
		assert.equal(readTransition(slide)?.kind, "cut");
	});

	it("removes one", () => {
		writeTransition(slide, { kind: "fade", speed: "med", direction: null });
		assert.equal(writeTransition(slide, null), true);
		assert.equal(readTransition(slide), null);
		assert.equal(writeTransition(slide, null), false);
	});

	it("sits where the schema puts it, before the timing tree", () => {
		addAnimation(slide, "5", "fade", "click");
		writeTransition(slide, { kind: "fade", speed: "med", direction: null });
		assert.deepEqual(childOrder(), ["cSld", "clrMapOvr", "transition", "timing"]);
	});

	it("defaults an unreadable speed to medium", () => {
		writeTransition(slide, { kind: "fade", speed: "med", direction: null });
		all("transition")[0].setAttribute("spd", "nonsense");
		assert.equal(readTransition(slide)?.speed, "med");
	});
});

describe("adding animations", () => {
	it("builds a timing tree on a slide that had none", () => {
		assert.equal(addAnimation(slide, "5", "fade", "click"), true);
		assert.equal(all("timing").length, 1);
		assert.equal(all("seq").length, 1);
		assert.deepEqual(readAnimations(slide), [
			{ shapeId: "5", effect: "fade", trigger: "click", index: 0 },
		]);
	});

	it("targets the shape it was given, everywhere it says so", () => {
		addAnimation(slide, "42", "fade", "click");
		const targets = all("spTgt").map((el) => el.getAttribute("spid"));
		assert.ok(targets.length >= 2, "an entrance sets visibility and then animates");
		assert.deepEqual(new Set(targets), new Set(["42"]));
	});

	it("writes each effect with its own preset", () => {
		addAnimation(slide, "1", "fade", "click");
		addAnimation(slide, "2", "wipe", "click");
		addAnimation(slide, "3", "flyIn", "click");
		assert.deepEqual(
			readAnimations(slide).map((a) => a.effect),
			["fade", "wipe", "flyIn"],
		);
	});

	it("moves a fly-in along both axes", () => {
		addAnimation(slide, "1", "flyIn", "click");
		const names = all("attrName").map((el) => el.textContent);
		assert.ok(names.includes("ppt_x"));
		assert.ok(names.includes("ppt_y"));
	});

	it("gives each click its own group", () => {
		addAnimation(slide, "1", "fade", "click");
		addAnimation(slide, "2", "fade", "click");
		const seqList = all("seq")[0].getElementsByTagName("*");
		void seqList;
		assert.deepEqual(
			readAnimations(slide).map((a) => a.trigger),
			["click", "click"],
		);
		// Two clicks, so two groups hanging off the main sequence.
		const mainList = all("cTn").find((el) => el.getAttribute("nodeType") === "mainSeq");
		const groups = Array.from(mainList?.firstElementChild?.children ?? []);
		assert.equal(groups.length, 2);
	});

	it("puts a with-previous effect alongside the one before it", () => {
		addAnimation(slide, "1", "fade", "click");
		addAnimation(slide, "2", "fade", "withPrev");
		const mainList = all("cTn").find((el) => el.getAttribute("nodeType") === "mainSeq");
		const groups = Array.from(mainList?.firstElementChild?.children ?? []);
		assert.equal(groups.length, 1, "with-previous does not start a new click");
		assert.deepEqual(
			readAnimations(slide).map((a) => a.trigger),
			["click", "withPrev"],
		);
	});

	it("puts an after-previous effect in the same click, but after it", () => {
		addAnimation(slide, "1", "fade", "click");
		addAnimation(slide, "2", "fade", "afterPrev");
		const mainList = all("cTn").find((el) => el.getAttribute("nodeType") === "mainSeq");
		const groups = Array.from(mainList?.firstElementChild?.children ?? []);
		assert.equal(groups.length, 1);
		const middles = Array.from(groups[0].firstElementChild?.lastElementChild?.children ?? []);
		assert.equal(middles.length, 2, "after-previous waits its turn within the click");
	});

	it("makes the first effect a click, whatever it was asked for", () => {
		// Nothing precedes it, so "with previous" would describe a wait on
		// something that does not exist.
		addAnimation(slide, "1", "fade", "withPrev");
		assert.deepEqual(readAnimations(slide), [
			{ shapeId: "1", effect: "fade", trigger: "click", index: 0 },
		]);
	});

	it("numbers every time node uniquely, in document order", () => {
		addAnimation(slide, "1", "fade", "click");
		addAnimation(slide, "2", "wipe", "withPrev");
		addAnimation(slide, "3", "flyIn", "afterPrev");
		const ids = all("cTn").map((el) => Number(el.getAttribute("id")));
		assert.deepEqual(
			ids,
			ids.map((_, i) => i + 1),
			"ids should run 1..n with no gaps or repeats",
		);
	});

	it("reports an effect it did not write as one it does not know", () => {
		addAnimation(slide, "1", "fade", "click");
		all("cTn")
			.find((el) => el.getAttribute("presetID") === "10")
			?.setAttribute("presetID", "53");
		assert.equal(readAnimations(slide)[0].effect, "other");
	});
});

describe("removing animations", () => {
	it("takes the timing tree away with the last effect", () => {
		addAnimation(slide, "5", "fade", "click");
		assert.equal(removeAnimationsFor(slide, "5"), 1);
		assert.equal(all("timing").length, 0);
		assert.deepEqual(childOrder(), ["cSld", "clrMapOvr"]);
	});

	it("leaves no empty wrapper behind", () => {
		addAnimation(slide, "1", "fade", "click");
		addAnimation(slide, "2", "fade", "click");
		removeAnimationsFor(slide, "1");
		assert.deepEqual(
			readAnimations(slide).map((a) => a.shapeId),
			["2"],
		);
		for (const par of all("par")) {
			assert.ok(par.firstElementChild, "a par with no children is invalid");
		}
		for (const list of all("childTnLst")) {
			assert.ok(list.firstElementChild, "an empty childTnLst should have been pruned");
		}
	});

	it("removes every effect on one shape at once", () => {
		addAnimation(slide, "1", "fade", "click");
		addAnimation(slide, "1", "wipe", "click");
		addAnimation(slide, "2", "fade", "click");
		assert.equal(removeAnimationsFor(slide, "1"), 2);
		assert.deepEqual(
			readAnimations(slide).map((a) => a.shapeId),
			["2"],
		);
	});

	it("does nothing for a shape with no animation", () => {
		addAnimation(slide, "1", "fade", "click");
		assert.equal(removeAnimationsFor(slide, "9"), 0);
		assert.equal(readAnimations(slide).length, 1);
	});

	it("removes one by position", () => {
		addAnimation(slide, "1", "fade", "click");
		addAnimation(slide, "2", "wipe", "click");
		assert.equal(removeAnimationAt(slide, 0), true);
		assert.deepEqual(
			readAnimations(slide).map((a) => a.shapeId),
			["2"],
		);
		assert.equal(removeAnimationAt(slide, 5), false);
	});

	it("renumbers what is left", () => {
		addAnimation(slide, "1", "fade", "click");
		addAnimation(slide, "2", "fade", "click");
		addAnimation(slide, "3", "fade", "click");
		removeAnimationAt(slide, 1);
		const ids = all("cTn").map((el) => Number(el.getAttribute("id")));
		assert.deepEqual(
			ids,
			ids.map((_, i) => i + 1),
		);
	});

	it("clears the lot", () => {
		addAnimation(slide, "1", "fade", "click");
		assert.equal(clearAnimations(slide), true);
		assert.equal(readAnimations(slide).length, 0);
		assert.equal(clearAnimations(slide), false);
	});
});

describe("reordering", () => {
	it("moves an effect later", () => {
		addAnimation(slide, "1", "fade", "click");
		addAnimation(slide, "2", "wipe", "click");
		addAnimation(slide, "3", "flyIn", "click");
		assert.equal(moveAnimation(slide, 0, 2), true);
		assert.deepEqual(
			readAnimations(slide).map((a) => a.shapeId),
			["2", "3", "1"],
		);
	});

	it("keeps each effect's own trigger when it moves", () => {
		addAnimation(slide, "1", "fade", "click");
		addAnimation(slide, "2", "wipe", "withPrev");
		addAnimation(slide, "3", "flyIn", "afterPrev");
		assert.equal(moveAnimation(slide, 2, 1), true);
		assert.deepEqual(readAnimations(slide), [
			{ shapeId: "1", effect: "fade", trigger: "click", index: 0 },
			{ shapeId: "3", effect: "flyIn", trigger: "afterPrev", index: 1 },
			{ shapeId: "2", effect: "wipe", trigger: "withPrev", index: 2 },
		]);
	});

	it("promotes an effect moved to the front, which no longer has a previous", () => {
		addAnimation(slide, "1", "fade", "click");
		addAnimation(slide, "2", "wipe", "withPrev");
		assert.equal(moveAnimation(slide, 1, 0), true);
		assert.deepEqual(readAnimations(slide), [
			{ shapeId: "2", effect: "wipe", trigger: "click", index: 0 },
			{ shapeId: "1", effect: "fade", trigger: "click", index: 1 },
		]);
	});

	it("leaves the tree valid after a move", () => {
		addAnimation(slide, "1", "fade", "click");
		addAnimation(slide, "2", "wipe", "click");
		moveAnimation(slide, 1, 0);
		const ids = all("cTn").map((el) => Number(el.getAttribute("id")));
		assert.deepEqual(
			ids,
			ids.map((_, i) => i + 1),
		);
		for (const par of all("par")) assert.ok(par.firstElementChild);
	});

	it("refuses a move that goes nowhere or off the end", () => {
		addAnimation(slide, "1", "fade", "click");
		addAnimation(slide, "2", "fade", "click");
		assert.equal(moveAnimation(slide, 0, 0), false);
		assert.equal(moveAnimation(slide, 0, 9), false);
		assert.equal(moveAnimation(slide, -1, 0), false);
	});

	it("will not rebuild a sequence holding an effect it cannot write", () => {
		addAnimation(slide, "1", "fade", "click");
		addAnimation(slide, "2", "fade", "click");
		all("cTn")
			.find((el) => el.getAttribute("presetID") === "10")
			?.setAttribute("presetID", "53");
		assert.equal(canReorder(slide), false);
		assert.equal(moveAnimation(slide, 0, 1), false);
		assert.equal(readAnimations(slide).length, 2, "and it leaves the slide alone");
	});

	it("has nothing to reorder below two effects", () => {
		assert.equal(canReorder(slide), false);
		addAnimation(slide, "1", "fade", "click");
		assert.equal(canReorder(slide), false);
		addAnimation(slide, "2", "fade", "click");
		assert.equal(canReorder(slide), true);
	});
});
