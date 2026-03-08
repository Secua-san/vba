import esbuild from "esbuild";
import { mkdir, rm } from "node:fs/promises";

await rm("dist", { force: true, recursive: true });
await mkdir("dist", { recursive: true });

await esbuild.build({
  bundle: true,
  entryPoints: ["src/extension.ts"],
  external: ["vscode"],
  format: "cjs",
  outfile: "dist/extension.js",
  platform: "node",
  sourcemap: true,
  target: "node20"
});
