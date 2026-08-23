/**
 * Bakes the theme and component galleries.
 *
 * Everything on /themes and /components comes from the package's own registry —
 * `docxcelerate/registry` for the catalog, `registry/` for the component
 * sources — resolved through the real framework and the real renderer. The site
 * describes nothing itself, so a theme whose colours change or a component whose
 * code changes updates the pages by being rebuilt.
 *
 * Output:
 *   public/demo/themes/<id>.html       one sample document, set in that theme
 *   public/demo/components/<id>.html   the component, resolved against its own
 *                                      preview data
 *   src/generated/registry.json        the catalog, plus the source of every
 *                                      file `dxcl add` would copy in
 */
import { build } from "esbuild";
import { docxcelerateEsbuildTransform } from "docxcelerate/transform";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { dirname, posix, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { extractPage, NODE_ONLY_STYLE, PAGE_ONLY_STYLE } from "./lib/preview-page.mjs";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..");

const TEMP_DIR = resolve(ROOT, ".registry-build");
const THEME_OUT = resolve(ROOT, "public/demo/themes");
const COMPONENT_OUT = resolve(ROOT, "public/demo/components");
const MANIFEST = resolve(ROOT, "src/generated/registry.json");

/** Where previews are addressed from in the browser. */
const THEME_BASE = "/demo/themes";
const COMPONENT_BASE = "/demo/components";

async function main() {
  await rm(TEMP_DIR, { recursive: true, force: true });
  // Wiped rather than overwritten, so a theme or component that left the
  // registry takes its preview with it instead of leaving a page behind.
  await rm(THEME_OUT, { recursive: true, force: true });
  await rm(COMPONENT_OUT, { recursive: true, force: true });
  await mkdir(TEMP_DIR, { recursive: true });
  await mkdir(THEME_OUT, { recursive: true });
  await mkdir(COMPONENT_OUT, { recursive: true });
  await mkdir(dirname(MANIFEST), { recursive: true });

  const {
    COMPONENT_CATEGORIES,
    COMPONENTS,
    THEMES,
    themeById,
  } = await import("docxcelerate");
  const { registryRoot } = await import("docxcelerate/registry/install");
  const { buildDocument } = await import("docxcelerate");
  const { renderDocumentWebsite } = await import("docxcelerate/renderer");

  const themes = [];

  for (const theme of THEMES) {
    const rendered = renderDocumentWebsite(
      { ...SAMPLE_DOCUMENT, title: `${theme.title} sample`, style: theme.style },
      { title: theme.title },
    );

    await writeFile(
      resolve(THEME_OUT, `${theme.id}.html`),
      extractPage(rendered, { title: theme.title, style: PAGE_ONLY_STYLE }),
      "utf8",
    );

    themes.push({
      id: theme.id,
      title: theme.title,
      summary: theme.summary,
      detail: theme.detail,
      category: theme.category,
      tags: [...theme.tags],
      fonts: [...theme.fonts],
      // The style itself, because it is what a reader is choosing between and
      // what the install writes. Shown as numbers rather than described.
      style: theme.style,
      previewUrl: posix.join(THEME_BASE, `${theme.id}.html`),
    });

    console.log(`registry: theme ${theme.id.padEnd(16)} ${theme.fonts.join(", ")}`);
  }

  const wrap = await documentTemplate();
  const components = [];

  for (const component of COMPONENTS) {
    const files = [];

    for (const file of component.files) {
      files.push({
        source: file.source,
        target: file.target,
        code: await readFile(resolve(registryRoot(), ...file.source.split("/")), "utf8"),
      });
    }

    const module = await loadComponent(component, registryRoot());
    const exported = component.exports.map((name) => {
      if (typeof module[name] !== "function") {
        throw new Error(
          `Component "${component.id}" says it exports ${name}, but ` +
            `${component.files[0].source} does not.`,
        );
      }

      return module[name];
    });

    // Placeholder mode, so a generated node shows what an author sees rather
    // than requiring an endpoint to build the docs.
    const document = await buildDocument(
      wrap(component.id, component.title, exported),
      component.previewData,
      { dynamicMode: "placeholder" },
    );

    // Components drawn against a theme are previewed in it; the rest get the
    // default, which is what a project has before it chooses one.
    const style = themeById(component.themeHint ?? "clean-minimal").style;
    const rendered = renderDocumentWebsite({ ...document, style }, { title: component.title });

    await writeFile(
      resolve(COMPONENT_OUT, `${component.id}.html`),
      extractPage(rendered, { title: component.title, style: NODE_ONLY_STYLE }),
      "utf8",
    );

    components.push({
      id: component.id,
      title: component.title,
      summary: component.summary,
      detail: component.detail,
      category: component.category,
      tags: [...component.tags],
      exports: [...component.exports],
      dataFields: component.dataFields,
      requires: component.requires,
      themeHint: component.themeHint ?? null,
      files,
      previewData: component.previewData,
      // What the component resolved to, so the page can show the nodes an
      // engine or a renderer would be handed.
      resolved: document.nodes,
      previewUrl: posix.join(COMPONENT_BASE, `${component.id}.html`),
    });

    console.log(
      `registry: component ${component.id.padEnd(16)} ${document.nodes.length} nodes`,
    );
  }

  await writeFile(
    MANIFEST,
    `${
      JSON.stringify(
        {
          themes,
          themeCategories: [...new Set(themes.map((theme) => theme.category))],
          components,
          componentCategories: COMPONENT_CATEGORIES,
        },
        null,
        2,
      )
    }\n`,
    "utf8",
  );

  console.log(
    `registry: ${themes.length} themes, ${components.length} components -> ` +
      `public/demo/, manifest -> src/generated/registry.json`,
  );

  await rm(TEMP_DIR, { recursive: true, force: true });
}

/**
 * Wraps a component's exports in a document, without JSX.
 *
 * This script is plain JavaScript, so it reaches for the same jsx() call a
 * compiled `<Document>` would make — the same trick the node gallery uses, for
 * the same reason: no build step between the registry and what the site shows.
 */
async function documentTemplate() {
  const { jsx, jsxs } = await import("docxcelerate/template/jsx-runtime");
  const { Document, template } = await import("docxcelerate/template");

  // jsxs rather than jsx, and no keys: a component's exports are siblings
  // somebody wrote out, not a list an expression produced. Keying them would
  // suffix every id — the page would show `salutation-0` for a node a project
  // would call `salutation`.
  return (id, title, components) =>
    template(
      jsxs(Document, {
        id,
        title,
        children: components.map((component) => jsx(component, {})),
      }),
    );
}

/**
 * Loads one component's source the way a document project would: bundled so its
 * relative imports resolve, with `docxcelerate` left external so it comes
 * through website/node_modules and the package's own exports map.
 */
async function loadComponent(component, registryDir) {
  const entry = resolve(TEMP_DIR, `${component.id}.mjs`);

  await build({
    entryPoints: [resolve(registryDir, ...component.files[0].source.split("/"))],
    outfile: entry,
    bundle: true,
    format: "esm",
    platform: "node",
    target: "node20",
    jsx: "automatic",
    jsxImportSource: "docxcelerate/template",
    packages: "external",
    plugins: [docxcelerateEsbuildTransform()],
    logLevel: "warning",
  });

  return await import(pathToFileURL(entry).href);
}

/**
 * One document, set in every theme in turn.
 *
 * Written as a model rather than as a template because a theme sample is about
 * type and colour, not about authoring: what it needs is a title, a heading, a
 * few paragraphs of ordinary length, and one placeholder to show how a theme
 * treats the parts of a document that are not prose.
 */
const SAMPLE_DOCUMENT = {
  schemaVersion: "docxcelerate.letter/v0",
  id: "theme-sample",
  title: "Membership renewal",
  nodes: [
    {
      id: "opening",
      kind: "paragraph",
      mode: "static",
      text:
        "Dear Ms Nkemelu, your Peak Anytime membership at Riverside Leisure " +
        "Centre renews on 1 October 2026. Nothing needs doing — this letter " +
        "sets out what changes and what does not.",
    },
    {
      id: "price",
      kind: "section",
      title: "What it costs",
      children: [
        {
          id: "price-change",
          kind: "paragraph",
          mode: "static",
          text:
            "Your membership is rising by 5.1%, from £468 to £492 a year. " +
            "That is £41 a month from 1 October 2026, taken on the same date " +
            "as it is now.",
        },
        {
          id: "price-hold",
          kind: "paragraph",
          mode: "static",
          text:
            "The price is then held for twelve months. We review prices once " +
            "a year and write to you before anything changes.",
        },
      ],
    },
    {
      id: "your-year",
      kind: "section",
      title: "Your year here",
      children: [
        {
          id: "visits",
          kind: "graph",
          mode: "static",
          graphType: "bar",
          caption: "Visits by month",
          data: { series: [11, 14, 9, 16, 18, 12] },
        },
        {
          id: "visits-note",
          kind: "paragraph",
          mode: "static",
          text:
            "You came in 80 times this year, most often on Tuesday evenings. " +
            "Swimming was 42% of that, strength work another third.",
        },
      ],
    },
  ],
};

main().catch((error) => {
  console.error("registry build failed:", error);
  process.exitCode = 1;
});
