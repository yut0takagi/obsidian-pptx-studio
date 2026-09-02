/**
 * Choosing what to paste.
 *
 * The clipboard usually offers the same picture more than one way, so the
 * cases that matter are the ones with several candidates and the ones with a
 * candidate that only looks like an image.
 */
import { strict as assert } from "node:assert";
import { describe, it } from "node:test";

import {
	chooseImage,
	clipboardImageName,
	extensionForImageType,
} from "../../src/edit/clipboardImage";

const file = (type: string, name?: string) => ({ type, name });

describe("extensionForImageType", () => {
	it("maps the types a clipboard actually offers", () => {
		assert.equal(extensionForImageType("image/png"), "png");
		assert.equal(extensionForImageType("image/jpeg"), "jpg");
		assert.equal(extensionForImageType("image/svg+xml"), "svg");
	});

	it("ignores parameters and case", () => {
		assert.equal(extensionForImageType("IMAGE/PNG; charset=binary"), "png");
	});

	it("refuses anything that is not an image it can store", () => {
		assert.equal(extensionForImageType("text/html"), null);
		assert.equal(extensionForImageType("image/tiff"), null);
		assert.equal(extensionForImageType(""), null);
	});
});

describe("chooseImage", () => {
	it("prefers PNG over the other representations offered alongside it", () => {
		const chosen = chooseImage([file("image/jpeg"), file("image/png"), file("image/webp")]);
		assert.equal(chosen?.file.type, "image/png");
		assert.equal(chosen?.extension, "png");
	});

	it("takes what it can when there is no PNG", () => {
		assert.equal(chooseImage([file("image/webp"), file("image/gif")])?.extension, "gif");
	});

	it("leaves SVG for last, since it renders unevenly elsewhere", () => {
		assert.equal(chooseImage([file("image/svg+xml"), file("image/bmp")])?.extension, "bmp");
	});

	it("skips the text a copied image is usually accompanied by", () => {
		const chosen = chooseImage([file("text/html"), file("text/plain"), file("image/png")]);
		assert.equal(chosen?.extension, "png");
	});

	it("returns null when the clipboard holds no image at all", () => {
		assert.equal(chooseImage([file("text/plain")]), null);
		assert.equal(chooseImage([]), null);
	});
});

describe("clipboardImageName", () => {
	it("keeps a name that already says what it is", () => {
		assert.equal(clipboardImageName("diagram.png", "png"), "diagram.png");
	});

	it("gives an extension to a name that has none", () => {
		assert.equal(clipboardImageName("diagram", "png"), "diagram.png");
	});

	it("names a screenshot, which arrives with nothing", () => {
		assert.equal(clipboardImageName(undefined, "png"), "clipboard.png");
		assert.equal(clipboardImageName("   ", "jpg"), "clipboard.jpg");
	});
});
