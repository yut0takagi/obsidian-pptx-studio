import esbuild from "esbuild";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = dirname(fileURLToPath(import.meta.url));

// jsdom loads data files relative to its own package at runtime, so it has to
// stay external rather than being inlined into the bundle.
await esbuild.build({
	entryPoints: ["scripts/smoke.ts"],
	bundle: true,
	platform: "node",
	format: "cjs",
	target: "node20",
	outfile: "scripts/smoke.cjs",
	external: ["jsdom"],
	plugins: [
		{
			name: "obsidian-shim",
			setup(build) {
				build.onResolve({ filter: /^obsidian$/ }, () => ({
					path: join(root, "scripts/obsidian-shim.ts"),
				}));
			},
		},
	],
	logLevel: "warning",
});
