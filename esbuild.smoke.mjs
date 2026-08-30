import esbuild from "esbuild";

await esbuild.build({
	entryPoints: ["scripts/smoke.ts"],
	bundle: true,
	platform: "node",
	format: "cjs",
	target: "node20",
	outfile: "scripts/smoke.cjs",
	logLevel: "warning",
});
