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
import { buildDocument } from "docxcelerate";
import { Document, template } from "docxcelerate/template";
import { jsx } from "docxcelerate/template/jsx-runtime";
import { loadRegistryModule } from "./registry_support.ts";

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

      // The node has to export what the catalog says it does, or the export
      // line the install appends points at nothing. A deriver beside it exports
      // a deriver, and is held to that below instead.
      if (!file.target.startsWith("nodes/")) continue;

      for (const name of component.exports) {
        assertStringIncludes(source, `export const ${name}`);
      }
    }
  }
});

test("every component compiles and loads the way a project's build would load it", async () => {
  // Through the transform, so a node whose control flow does not compile fails
  // here rather than on somebody's first preview.
  for (const component of COMPONENTS) {
    const module = await loadRegistryModule<Record<string, unknown>>(component.files[0].source);

    for (const name of component.exports) {
      assertEquals(`${component.id}.${name}: ${typeof module[name]}`, `${component.id}.${name}: function`);
    }
  }
});

test("every component is laid out as a node, at most one deriver, and its arrangements", async () => {
  // One layout, so the node's `../derivers/` and `../arrangements/` imports
  // resolve both in the registry and once installed, and so the installer's
  // follow-up can name the deriver without reading it. Collected, so a failure
  // names every entry.
  const wrong: string[] = [];

  for (const component of COMPONENTS) {
    const { id } = component;
    const nodeFile = { source: `components/${id}/${id}.node.tsx`, target: `nodes/${id}.node.tsx` };
    const deriverFile = { source: `components/derivers/${id}.deriver.ts`, target: `derivers/${id}.deriver.ts` };
    const arranged = (file: string) => ({
      source: `components/arrangements/${id}/${file}`,
      target: `arrangements/${id}/${file}`,
    });
    // Every arrangement the catalog lists has its file, and nothing else is there.
    const arrangementFiles = component.arrangements
      ? [arranged("lines.ts"), ...component.arrangements.map((a) => arranged(`${a.id}.arrangement.tsx`))]
      : [];
    const [first, ...others] = component.files;
    const hasDeriver = JSON.stringify(others[0]) === JSON.stringify(deriverFile);
    const rest = hasDeriver ? others.slice(1) : others;

    if (JSON.stringify(first) !== JSON.stringify(nodeFile)) wrong.push(`${id}: first file is not ${nodeFile.source}`);
    if (JSON.stringify(rest) !== JSON.stringify(arrangementFiles)) {
      wrong.push(`${id}: after the node and deriver, files are not ${JSON.stringify(arrangementFiles)}`);
    }

    if (!hasDeriver) continue;

    const source = await readFile(join(registryRoot(), ...deriverFile.source.split("/")), "utf8");
    const name = id.replace(/-([a-z0-9])/g, (_, c: string) => c.toUpperCase());

    if (!source.includes("export default deriver(")) wrong.push(`${id}: deriver is not a default deriver() export`);
    if (!source.includes(`name: "${name}"`)) wrong.push(`${id}: deriver is not named ${name}`);
  }

  assertEquals(wrong, []);
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

// What a catalogued arrangement promises: the component takes it, and it places
// things somewhere the others do not. A component that ignored the prop would
// pass a build and show the same preview under every caption.
test("every arrangement a component lists is one it builds, and each is its own", async () => {
  for (const component of COMPONENTS) {
    if (!component.arrangements) continue;

    const ids = component.arrangements.map((arrangement) => arrangement.id);
    assertEquals(`${component.id}: ${new Set(ids).size >= 2}`, `${component.id}: true`);
    assertEquals(new Set(ids).size, ids.length);

    const module = await loadRegistryModule<Record<string, unknown>>(component.files[0].source);
    const build = (props: Record<string, unknown>) =>
      buildDocument(
        template(
          jsx(Document, {
            id: component.id,
            title: component.title,
            children: jsx(module[component.exports[0]] as never, props),
          }),
        ),
        component.previewData,
        { dynamicMode: "placeholder" },
      );
    const shapes = new Set<string>();

    for (const id of ids) {
      shapes.add(JSON.stringify((await build({ arrangement: id })).nodes));
    }

    assertEquals(`${component.id}: ${shapes.size} distinct`, `${component.id}: ${ids.length} distinct`);
    // No prop is the first arrangement, as the catalog lists them.
    assertEquals(
      JSON.stringify((await build({})).nodes),
      JSON.stringify((await build({ arrangement: ids[0] })).nodes),
    );
    await assertRejects(() => build({ arrangement: "no-such-arrangement" }), Error, "arrangement");
  }
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

test("installing a component with arrangements lands them where its node imports them from", async () => {
  const projectDir = await documentProject();

  const result = await installRegistryEntry({ ref: "invoice-document-type", projectDir });
  const installed = await readFile(join(projectDir, "nodes", "invoice-document-type.node.tsx"), "utf8");

  for (const file of ["lines.ts", "stacked.arrangement.tsx", "table.arrangement.tsx", "split.arrangement.tsx"]) {
    assertStringIncludes(result.files.join("\n"), `arrangements/invoice-document-type/${file}`);
    assertStringIncludes(installed, `"../arrangements/invoice-document-type/${file}"`);
    await readFile(join(projectDir, "arrangements", "invoice-document-type", file), "utf8");
  }
  assertStringIncludes(result.followUp.join("\n"), 'arrangement=: "stacked", "table", "split"');
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
