import { defineConfig } from "tsdown";

// The agent skill's CLI: one self-contained ES module (library and dependencies
// inlined), so the skill runs with nothing but Node 20+.
export default defineConfig({
  entry: { docxtor: "cli/docxtor.ts" },
  outDir: "plugin/skills/docxtor/scripts",
  format: ["esm"],
  platform: "node",
  target: "node20",
  noExternal: [/.*/],
  minify: true,
  dts: false,
  sourcemap: false,
  clean: false,
  treeshake: true,
  outExtensions: () => ({ js: ".mjs" }),
  banner: { js: "#!/usr/bin/env node\n// Built from cli/ by `pnpm build:skill`. Do not edit." },
});
