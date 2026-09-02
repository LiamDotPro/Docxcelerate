import { test } from "node:test";
import { readFile, stat } from "node:fs/promises";
import { assertEquals } from "./assert.ts";
import { version } from "docxcelerate";

/**
 * The two package manifests, and whether they still describe the same package.
 *
 * npm reads `package.json` and JSR reads `deno.json`, and the export maps in
 * them are maintained by hand. Nothing had been comparing them, so they drifted:
 * `./transform` shipped to one registry and not the other, and a `./preview`
 * subpath pointed at a file the build had stopped emitting — an import that
 * threw for anyone who tried it, and that no test could have caught, because
 * the failure is in the manifest rather than in any code.
 *
 * These run against the built `dist/`, which is what `package.json` points at.
 * `npm test` builds first, so they are testing the thing that would be
 * published rather than the sources it came from.
 */

const root = new URL("../", import.meta.url);

async function readJson(name: string): Promise<Record<string, unknown>> {
  return JSON.parse(await readFile(new URL(name, root), "utf8"));
}

async function exists(path: URL): Promise<boolean> {
  try {
    await stat(path);
    return true;
  } catch {
    return false;
  }
}

test("both registries publish the same set of entrypoints", async () => {
  const npm = await readJson("package.json");
  const jsr = await readJson("deno.json");

  const npmSubpaths = Object.keys(npm.exports as object).sort();
  const jsrSubpaths = Object.keys(jsr.exports as object).sort();

  assertEquals(npmSubpaths, jsrSubpaths);
});

test("every entrypoint npm publishes exists once the package is built", async () => {
  const npm = await readJson("package.json");
  const entries = Object.entries(npm.exports as Record<string, Record<string, string>>);
  const missing: string[] = [];

  for (const [subpath, target] of entries) {
    for (const file of [target.types, target.default]) {
      if (file && !(await exists(new URL(file, root)))) {
        missing.push(`${subpath} -> ${file}`);
      }
    }
  }

  assertEquals(missing, []);
});

test("an entrypoint npm publishes is built from the source JSR publishes", async () => {
  const npm = await readJson("package.json");
  const jsr = await readJson("deno.json");
  const npmExports = npm.exports as Record<string, Record<string, string>>;
  const jsrExports = jsr.exports as Record<string, string>;
  const mismatched: string[] = [];

  for (const [subpath, source] of Object.entries(jsrExports)) {
    // ./src/project/scaffold.ts is built to ./dist/project/scaffold.js, so one
    // is the other with the roots and the extension swapped. A subpath whose
    // two manifests name unrelated modules is the drift this catches.
    const built = source.replace("./src/", "./dist/").replace(/\.ts$/, ".js");

    if (npmExports[subpath]?.default !== built) {
      mismatched.push(`${subpath}: jsr ${source}, npm ${npmExports[subpath]?.default}`);
    }
  }

  assertEquals(mismatched, []);
});

test("the version is the same number in all three places that carry it", async () => {
  const npm = await readJson("package.json");
  const jsr = await readJson("deno.json");

  assertEquals(npm.version, version);
  assertEquals(jsr.version, version);
});
