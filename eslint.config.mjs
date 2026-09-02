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
]);
