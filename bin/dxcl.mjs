#!/usr/bin/env node
import { existsSync } from "node:fs";
import { dirname, join } from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const packageRoot = dirname(dirname(fileURLToPath(import.meta.url)));
const cliEntry = join(packageRoot, "dist", "project", "scaffold_cli.js");

if (!existsSync(cliEntry)) {
  console.error(
    "Docxcelerate is not built yet. Run `npm run build` in the package directory first.",
  );
  process.exit(1);
}

process.on("unhandledRejection", (error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});

await import(`file://${cliEntry.replaceAll("\\", "/")}`);
