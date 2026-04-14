import { defineConfig } from "tsdown";

export default defineConfig({
  entry: ["src/index.ts"],
  format: ["esm", "cjs"],
  publint:true,
  attw: true,
  minify: true,
  dts: true,
  sourcemap: false,
  clean: true,
  treeshake: true,
  target: "node20",
});
