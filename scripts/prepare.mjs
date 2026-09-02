/**
 * The `prepare` lifecycle hook: build the package, if it can be built.
 *
 * `prepare` exists so consumers installing this package straight from git get
 * a compiled dist/. npm installs devDependencies before running it there, so
 * tsc is present and the build runs.
 *
 * It also fires in a case that has no business building: installing something
 * that merely links this directory — website/ depends on `file:..` — where npm
 * runs the hook without having installed this package's own devDependencies.
 * Failing there would fail the whole install. Skip instead; whoever needs
 * dist/ builds it themselves (website/scripts/build-lib.mjs does).
 */
import { execSync } from "node:child_process";
import { existsSync } from "node:fs";
import { dirname, resolve } from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");

if (!existsSync(resolve(ROOT, "node_modules/typescript/package.json"))) {
  console.log("prepare: typescript is not installed — skipping the build");
  process.exit(0);
}

execSync("npm run build", { cwd: ROOT, stdio: "inherit" });
