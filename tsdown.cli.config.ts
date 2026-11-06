import { defineConfig } from "tsdown";

// The `docxtor` command in the npm package (package.json "bin"). Same source as the
// skill's bundled CLI (tsdown.skill.config.ts), but the runtime dependencies stay
// external: npm installs them alongside.
export default defineConfig({
  entry: { cli: "cli/docxtor.ts" },
  outDir: "dist",
  format: ["esm"],
  platform: "node",
  target: "node20",
  minify: true,
  dts: false,
  sourcemap: false,
  clean: false,
  treeshake: true,
  outExtensions: () => ({ js: ".mjs" }),
  banner: { js: "#!/usr/bin/env node" },
});
