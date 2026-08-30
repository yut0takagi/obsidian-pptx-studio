import esbuild from "esbuild";

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
	logLevel: "warning",
});
