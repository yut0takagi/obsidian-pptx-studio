/**
 * Keep the three files that carry a version number in step.
 *
 * npm's `version` lifecycle runs this after package.json has been bumped and
 * before the release commit is made, so `npm version patch` is enough to carry
 * manifest.json and versions.json along with it. The release workflow checks
 * the same invariants against the tag; this is what keeps them from being
 * broken in the first place.
 *
 *   npm version patch                     # or minor / major
 *   node scripts/version-bump.mjs 0.2.0   # to finish a bump made by hand
 */
import { readFileSync, writeFileSync } from "node:fs";

const target = process.argv[2] ?? process.env.npm_package_version;
if (!target) {
	throw new Error("no version given: run this through `npm version`, or pass one as an argument");
}
if (!/^\d+\.\d+\.\d+$/.test(target)) {
	throw new Error(`not a release version: ${target}`);
}

/** Rewrite JSON in place, keeping the two-space style the files are checked in with. */
function edit(path, fn) {
	const json = JSON.parse(readFileSync(path, "utf8"));
	fn(json);
	writeFileSync(path, `${JSON.stringify(json, null, 2)}\n`);
}

// package.json is npm's to move, because the lockfile has to move with it.
const pkgVersion = JSON.parse(readFileSync("package.json", "utf8")).version;
if (pkgVersion !== target) {
	throw new Error(
		`package.json is at ${pkgVersion}, not ${target}: bump it with \`npm version\`, which updates the lockfile too`,
	);
}

let minAppVersion;
edit("manifest.json", (manifest) => {
	manifest.version = target;
	minAppVersion = manifest.minAppVersion;
});

// Obsidian reads versions.json to work out which build an older vault is still
// allowed to update to, so every released version needs an entry of its own.
edit("versions.json", (versions) => {
	versions[target] = minAppVersion;
});

console.log(`${target} (minAppVersion ${minAppVersion})`);
