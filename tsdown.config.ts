import { defineConfig } from "tsdown";

export default defineConfig({
  entry: ["src/index.ts"],
  format: ["esm", "cjs"],
  publint: false, // run by `pnpm lint` once the CLI (tsdown.cli.config.ts) is built too
  attw: true,
  minify: true,
  dts: true,
  sourcemap: false,
  clean: true,
  treeshake: true,
  target: "node20",
});
