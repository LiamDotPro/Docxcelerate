import { test } from "node:test";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { COMPONENTS, REGISTRY, REGISTRY_THEMES, findRegistryEntry } from "docxcelerate/registry";
import {
  installRegistryEntry,
  registryRoot,
  resolveInstallOrder,
} from "docxcelerate/registry/install";

/**
 * The registry, and what `dxcl add` leaves behind.
 *
 * An install is somebody's project being written into, so the tests care about
 * the two ways that goes wrong: putting a file somewhere that is not a project,
 * and overwriting something they had edited. Both are errors here rather than
 * surprises later.
 */

test("every registry id is unique across themes and components", () => {
  const ids = REGISTRY.map((entry) => entry.id);

  assertEquals(new Set(ids).size, ids.length);
});

test("every component names a source file that exists", async () => {
  for (const component of COMPONENTS) {
    for (const file of component.files) {
      const source = await readFile(join(registryRoot(), ...file.source.split("/")), "utf8");

      // The file has to export what the catalog says it does, or the export
      // line the install appends points at nothing.
      for (const name of component.exports) {
        assertStringIncludes(source, `export const ${name}`);
      }
    }
  }
});

test("every field a component documents is one its preview data actually has", () => {
  // The two halves of an entry drift in one direction: a component grows a
  // field, the catalog gains a row for it, and the preview data nobody has to
  // run stays as it was. What a reader then copies out of the site is data the
  // component reads a missing value from — which is not a crash, it is a
  // sentence with `undefined` in it.
  //
  // Collected rather than asserted one at a time, so a failure names every
  // entry that has drifted instead of only the first.
  const missing = COMPONENTS.flatMap((component) =>
    [...new Set(component.dataFields.map((field) => field.path.split(".")[0]))]
      .filter((root) => !(root in component.previewData))
      .map((root) => `${component.id}: previewData has no ${root}`)
  );

  assertEquals(missing, []);
});

test("a component that draws a chart says so in its tags", async () => {
  // Read out of the source rather than assumed from the id, so a component
  // that grows a chart later is held to this too. Tags are how the catalog is
  // filtered, and a chart component nobody searching for a chart can find is
  // a component that may as well not be in the registry.
  const untagged: string[] = [];

  for (const component of COMPONENTS) {
    for (const file of component.files) {
      const source = await readFile(join(registryRoot(), ...file.source.split("/")), "utf8");

      if (source.includes("<Graph") && !component.tags.includes("chart")) {
        untagged.push(component.id);
      }
    }
  }

  assertEquals(untagged, []);
});

test("installing a component copies it in and re-exports it", async () => {
  const projectDir = await documentProject();

  const result = await installRegistryEntry({ ref: "letterhead", projectDir });

  const installed = await readFile(join(projectDir, "nodes", "letterhead.node.tsx"), "utf8");
  const index = await readFile(join(projectDir, "nodes", "index.ts"), "utf8");

  assertEquals(result.kind, "component");
  assertEquals(result.files, ["nodes/letterhead.node.tsx", "nodes/index.ts"]);
  assertStringIncludes(installed, "export const Letterhead");
  assertStringIncludes(index, `export { Letterhead } from "./letterhead.node.tsx";`);
  // The fields it reads come back as work for a person, because nothing here
  // edits somebody's types.ts.
  assertStringIncludes(result.followUp.join("\n"), "sender.name");
});

test("installing the same component twice does not double the export", async () => {
  const projectDir = await documentProject();

  await installRegistryEntry({ ref: "letterhead", projectDir });
  await installRegistryEntry({ ref: "letterhead", projectDir, force: true });

  const index = await readFile(join(projectDir, "nodes", "index.ts"), "utf8");
  const exports = index.split("\n").filter((line) => line.includes("Letterhead"));

  assertEquals(exports.length, 1);
});

test("a component will not silently overwrite a file already there", async () => {
  const projectDir = await documentProject();
  await installRegistryEntry({ ref: "letterhead", projectDir });

  await assertRejects(
    () => installRegistryEntry({ ref: "letterhead", projectDir }),
    Error,
    "--force",
  );
});

test("installing a theme writes the project style", async () => {
  const projectDir = await documentProject();

  const result = await installRegistryEntry({ ref: "slate-report", projectDir });
  const style = await readFile(join(projectDir, "document-style.ts"), "utf8");

  assertEquals(result.kind, "theme");
  assertEquals(result.files, ["document-style.ts"]);
  assertStringIncludes(style, `import { slateReportTheme, themeStyle } from "docxcelerate/themes";`);
  assertStringIncludes(style, "export const documentStyle: DocumentStyle");
  // Which fonts it needs is the thing a person has to check before printing a
  // thousand of these, so the install says it rather than assuming.
  assertStringIncludes(result.followUp.join("\n"), "Calibri");
});

test("a second theme replaces the first, but not a style written by hand", async () => {
  const projectDir = await documentProject();
  await installRegistryEntry({ ref: "slate-report", projectDir });
  await installRegistryEntry({ ref: "warm-letter", projectDir });

  const style = await readFile(join(projectDir, "document-style.ts"), "utf8");
  assertStringIncludes(style, "warmLetterTheme");

  await writeFile(
    join(projectDir, "document-style.ts"),
    "export const documentStyle = { mine: true };\n",
    "utf8",
  );

  await assertRejects(
    () => installRegistryEntry({ ref: "legal-serif", projectDir }),
    Error,
    "--force",
  );
});

test("installing anywhere that is not a document project is an error", async () => {
  const projectDir = await mkdtemp(join(tmpdir(), "docxcelerate-not-a-project-"));

  await assertRejects(
    () => installRegistryEntry({ ref: "letterhead", projectDir }),
    Error,
    "document.project.ts",
  );
});

test("an id nothing carries lists what the registry does carry", async () => {
  const projectDir = await documentProject();

  await assertRejects(
    () => installRegistryEntry({ ref: "letterheed", projectDir }),
    Error,
    "slate-report",
  );
});

test("a kind prefix narrows the lookup", () => {
  assertEquals(findRegistryEntry("theme:slate-report")?.kind, "theme");
  assertEquals(findRegistryEntry("component:slate-report"), undefined);
  assertEquals(findRegistryEntry("component:letterhead")?.kind, "component");
});

test("install order puts requirements before what required them", () => {
  const order = resolveInstallOrder(["letterhead", "letterhead", "slate-report"]);

  assertEquals(order, ["letterhead", "slate-report"]);
  // Every shipped entry resolves, which is the cheapest way to catch a
  // requirement naming something that has been renamed or removed.
  assertEquals(
    resolveInstallOrder(REGISTRY.map((entry) => entry.id)).length,
    REGISTRY.length,
  );
});

test("every theme in the registry is one of the shipped themes", () => {
  for (const entry of REGISTRY_THEMES) {
    assertEquals(entry.theme.id, entry.id);
  }
});

/**
 * The smallest thing an install will accept: a directory with a
 * `document.project.ts` in it. Scaffolding a real project would work too, and
 * would tie these tests to what the scaffold happens to write.
 */
async function documentProject(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), "docxcelerate-project-"));

  await writeFile(join(dir, "document.project.ts"), "export default {};\n", "utf8");
  await writeFile(
    join(dir, "document-style.ts"),
    `import { cleanMinimalDocumentStyle } from "docxcelerate/document";\n` +
      `export const documentStyle = cleanMinimalDocumentStyle;\n`,
    "utf8",
  );

  return dir;
}
