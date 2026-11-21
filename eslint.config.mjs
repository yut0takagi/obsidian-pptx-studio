import { defineConfig, globalIgnores } from "eslint/config";
import obsidianmd from "eslint-plugin-obsidianmd";

/*
 * Only src is linted. It is the code that ships into Obsidian, and the rules
 * that matter here — no Node built-ins, createEl over innerHTML — are about
 * that runtime. The tests and build scripts run under Node, where the same
 * rules are noise, and tsc already typechecks them.
 */
export default defineConfig([
	globalIgnores(["main.js", "tests/", "scripts/", "*.mjs"]),
	...obsidianmd.configs.recommended,
	{
		languageOptions: {
			parserOptions: { projectService: true },
		},
	},
	/*
	 * The slide renderer styles everything inline, and it has to. PNG and PDF
	 * export serialise the rendered slide into an SVG <foreignObject>, where
	 * styles.css does not reach, so the fixed values this rule would move into
	 * classes are the ones an exported slide needs for its layout.
	 */
	{
		files: ["src/render/renderSlide.ts", "src/render/renderChart.ts"],
		rules: { "obsidianmd/no-static-styles-assignment": "off" },
	},
]);
