/**
 * Bundle the unit tests for Node's test runner.
 *
 * The sources import each other without file extensions, which Node's own
 * resolver will not follow, so the tests go through esbuild the same way the
 * smoke test does rather than being run from TypeScript directly.
 */
import esbuild from "esbuild";
import { readdirSync, rmSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = dirname(fileURLToPath(import.meta.url));
const unitDir = join(root, "tests/unit");
const outDir = join(root, "tests/.build");

const entryPoints = readdirSync(unitDir)
	.filter((name) => name.endsWith(".test.ts"))
	.sort()
	.map((name) => join(unitDir, name));

if (entryPoints.length === 0) {
	console.error("No test files found in tests/unit.");
	process.exit(1);
}

// A bundle outlives the test file it was built from, and a stale one keeps
// running and passing long after its source is gone — which reads as coverage
// that no longer exists. The directory is rebuilt from nothing every time.
rmSync(outDir, { recursive: true, force: true });

await esbuild.build({
	entryPoints,
	bundle: true,
	platform: "node",
	format: "cjs",
	target: "node20",
	outdir: outDir,
	outExtension: { ".js": ".cjs" },
	sourcemap: "inline",
	// jsdom loads data files relative to its own package at runtime, so it has to
	// stay external rather than being inlined into the bundle.
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
