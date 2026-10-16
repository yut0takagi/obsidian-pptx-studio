/**
 * Reading a slide's transition for playback.
 *
 * The editor's own reader (src/ooxml/animation.ts) only needs to recognise the
 * handful of transitions it can write. Playback has to cope with whatever a
 * deck carries, and most decks saved by a current PowerPoint wrap the
 * transition in `mc:AlternateContent`: a `p14` version with the real duration
 * and, for the newer effects, an effect only PowerPoint 2010 and later know,
 * next to a plain fallback. Reading only a direct `p:transition` child finds
 * nothing on those slides at all, which is why these cases are mostly about
 * that wrapper.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import { installDomParser } from "./dom";

installDomParser();

import { parseSlideTransition } from "../../src/pptx/transition";

const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";

function slide(inner: string): Element {
	return new DOMParser().parseFromString(
		`<p:sld xmlns:p="${P}" xmlns:mc="${MC}" xmlns:p14="${P14}"><p:cSld/>${inner}</p:sld>`,
		"application/xml",
	).documentElement;
}

describe("parseSlideTransition", () => {
	it("finds nothing on a slide without a transition", () => {
		assert.equal(parseSlideTransition(slide("")), null);
		assert.equal(parseSlideTransition(null), null);
	});

	it("reads a plain transition and its speed", () => {
		const t = parseSlideTransition(slide('<p:transition spd="fast"><p:fade/></p:transition>'));
		assert.equal(t?.effect, "fade");
		assert.equal(t?.duration, 500);
	});

	it("gives each named speed its own duration", () => {
		const of = (spd: string) =>
			parseSlideTransition(slide(`<p:transition spd="${spd}"><p:fade/></p:transition>`))?.duration;
		assert.equal(of("fast"), 500);
		assert.equal(of("med"), 750);
		assert.equal(of("slow"), 1000);
	});

	it("prefers an explicit p14:dur to the named speed", () => {
		const t = parseSlideTransition(
			slide('<p:transition spd="slow" p14:dur="1600"><p:push dir="u"/></p:transition>'),
		);
		assert.equal(t?.duration, 1600);
		assert.equal(t?.effect, "push");
		assert.equal(t?.direction, "u");
	});

	it("reads the transition PowerPoint wraps in mc:AlternateContent", () => {
		const t = parseSlideTransition(
			slide(
				"<mc:AlternateContent>" +
					'<mc:Choice Requires="p14"><p:transition spd="slow" p14:dur="2000"><p:wipe dir="r"/></p:transition></mc:Choice>' +
					'<mc:Fallback><p:transition spd="slow"><p:wipe dir="r"/></p:transition></mc:Fallback>' +
					"</mc:AlternateContent>",
			),
		);
		assert.equal(t?.effect, "wipe");
		assert.equal(t?.direction, "r");
		assert.equal(t?.duration, 2000);
	});

	it("falls back to the plain effect when the p14 one is beyond it, keeping the duration", () => {
		const t = parseSlideTransition(
			slide(
				"<mc:AlternateContent>" +
					'<mc:Choice Requires="p14"><p:transition spd="slow" p14:dur="1250"><p14:vortex dir="r"/></p:transition></mc:Choice>' +
					'<mc:Fallback><p:transition spd="slow"><p:push dir="d"/></p:transition></mc:Fallback>' +
					"</mc:AlternateContent>",
			),
		);
		assert.equal(t?.effect, "push");
		assert.equal(t?.direction, "d");
		assert.equal(t?.duration, 1250);
	});

	it("still plays something for an effect it has no drawing for", () => {
		const t = parseSlideTransition(slide('<p:transition><p:checker dir="vert"/></p:transition>'));
		assert.equal(t?.effect, "fade");
	});

	it("gives a travelling effect its default direction", () => {
		for (const name of ["push", "wipe", "cover", "pull"]) {
			const t = parseSlideTransition(slide(`<p:transition><p:${name}/></p:transition>`));
			assert.equal(t?.effect, name);
			assert.equal(t?.direction, "l", name);
		}
	});

	it("reads split's orientation and direction, with their defaults", () => {
		const given = parseSlideTransition(
			slide('<p:transition><p:split orient="vert" dir="in"/></p:transition>'),
		);
		assert.equal(given?.effect, "split");
		assert.equal(given?.orientation, "vert");
		assert.equal(given?.direction, "in");
		const bare = parseSlideTransition(slide("<p:transition><p:split/></p:transition>"));
		assert.equal(bare?.orientation, "horz");
		assert.equal(bare?.direction, "out");
	});

	it("reads a fade through black", () => {
		const t = parseSlideTransition(slide('<p:transition><p:fade thruBlk="1"/></p:transition>'));
		assert.equal(t?.throughBlack, true);
		const plain = parseSlideTransition(slide("<p:transition><p:fade/></p:transition>"));
		assert.equal(plain?.throughBlack, false);
	});

	it("reads an automatic advance, and treats a transition with no effect as a cut", () => {
		const t = parseSlideTransition(slide('<p:transition advTm="3000"/>'));
		assert.equal(t?.effect, "cut");
		assert.equal(t?.advanceAfter, 3000);
		const manual = parseSlideTransition(slide("<p:transition><p:fade/></p:transition>"));
		assert.equal(manual?.advanceAfter, null);
	});

	it("does not invent a duration for a cut", () => {
		const t = parseSlideTransition(slide('<p:transition spd="slow"><p:cut/></p:transition>'));
		assert.equal(t?.effect, "cut");
		assert.equal(t?.duration, 0);
	});
});
