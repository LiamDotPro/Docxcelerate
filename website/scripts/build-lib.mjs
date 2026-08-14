/**
 * Compiles the parent package before anything else in the site build.
 *
 * The site is a real consumer: every prep script imports `docxcelerate`, which
 * resolves through website/node_modules -> the parent's exports map -> its
 * dist/. That dist/ is gitignored, so a fresh checkout (CI, a clone, a deploy)
 * has no library to import until tsc has run. Building it here means the site
 * always renders against the current source, not a stale local build.
 *
 * The parent's own dependencies are installed if they are missing. Installing
 * the site links the parent but installs neither what it needs to build itself
 * (typescript) nor what it needs at runtime (docx) — npm treats a `file:`
 * link as already installed, and Node resolves the parent's imports from the
 * parent's node_modules, not the site's.
 *
 * Everything here goes through `npm`, because on Deno Deploy `npm` is a shim
 * over Deno rather than real npm. Bare `npm install` and `npm run <script>`
 * are the two forms that shim reliably; `npm ci` and flags like --no-audit are
 * not, so they are avoided even though a deploy would otherwise want them.
 */
import { execSync } from "node:child_process";
import { existsSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url));
const PACKAGE_ROOT = resolve(HERE, "../..");

// Run through a shell: on Windows npm is a .cmd, which Node refuses to spawn
// directly. Every command below is a fixed literal, so there is nothing to
// interpolate or escape.
function npm(command) {
  execSync(`npm ${command}`, { cwd: PACKAGE_ROOT, stdio: "inherit" });
}

function installed(name) {
  return existsSync(resolve(PACKAGE_ROOT, `node_modules/${name}/package.json`));
}

// typescript builds dist/; docx is imported by what dist/ exports, so the prep
// scripts that follow load it. Either one missing means the parent has never
// been installed here.
if (!installed("typescript") || !installed("docx")) {
  console.log("lib: installing parent package dependencies (npm install)");
  npm("install");
}

npm("run build");

const ENTRY = resolve(PACKAGE_ROOT, "dist/public.js");
if (!existsSync(ENTRY)) {
  throw new Error(
    `lib: the build reported success but ${ENTRY} does not exist. ` +
      `The site imports docxcelerate through the parent's exports map, so ` +
      `every prep script after this one would fail to resolve it.`,
  );
}

console.log("lib: docxcelerate built -> dist/");
