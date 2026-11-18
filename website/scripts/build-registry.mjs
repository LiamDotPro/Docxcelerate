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
 *   public/demo/components/<id>--<value>.html
 *                                      the same, once per value of its first
 *                                      string-union field (see variantAxis)
 *   src/generated/registry.json        the catalog, plus the source of every
 *                                      file `dxcl add` would copy in
 */
import { build } from "esbuild";
import { docxcelerateEsbuildTransform } from "docxcelerate/transform";
import { createHash } from "node:crypto";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { dirname, posix, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { NODE_ONLY_STYLE, PAGE_ONLY_STYLE, renderDocxPage } from "./lib/docx-page.mjs";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..");

const TEMP_DIR = resolve(ROOT, ".registry-build");
const CACHE_DIR = resolve(ROOT, ".registry-cache");
const THEME_OUT = resolve(ROOT, "public/demo/themes");
const COMPONENT_OUT = resolve(ROOT, "public/demo/components");
const MANIFEST = resolve(ROOT, "src/generated/registry.json");

/**
 * Bumped when this script changes what it produces.
 *
 * The cache is keyed on the inputs — a component's source, its preview data,
 * the framework version. None of those change when the *script* starts writing
 * a different shape of entry, so this is the part that has to be moved by hand.
 * Getting it wrong serves a stale manifest, so move it whenever the output
 * shape changes.
 */
const CACHE_VERSION = 3;

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
  // Not wiped: it is the whole point that it survives a build.
  await mkdir(CACHE_DIR, { recursive: true });

  const {
    COMPONENT_CATEGORIES,
    COMPONENTS,
    THEMES,
    themeById,
  } = await import("docxcelerate");
  const { registryRoot } = await import("docxcelerate/registry/install");
  const { buildDocument } = await import("docxcelerate");

  const framework = await frameworkVersion();
  const themes = [];
  let reused = 0;

  for (const theme of THEMES) {
    const key = cacheKey({
      kind: "theme",
      id: theme.id,
      style: theme.style,
      sample: SAMPLE_DOCUMENT,
      framework,
    });
    const cached = await readCache(key);
    const page = cached?.page ?? await renderDocxPage(
      { ...SAMPLE_DOCUMENT, title: `${theme.title} sample`, style: theme.style },
      { title: theme.title, style: PAGE_ONLY_STYLE },
    );

    if (cached) {
      reused += 1;
    } else {
      await writeCache(key, { page });
    }

    await writeFile(resolve(THEME_OUT, `${theme.id}.html`), page, "utf8");

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

    console.log(
      `registry: theme ${theme.id.padEnd(16)} ${theme.fonts.join(", ")}` +
        `${cached ? " (cached)" : ""}`,
    );
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

    // Everything the preview is built from. Bundling and rendering a component
    // is the expensive half of this script — an esbuild pass each — and none of
    // it has to happen again while the inputs are the same.
    const key = cacheKey({
      kind: "component",
      id: component.id,
      files: files.map((file) => file.code),
      previewData: component.previewData,
      exports: component.exports,
      themeHint: component.themeHint ?? null,
      axis: variantAxis(component),
      framework,
    });
    const cached = await readCache(key);
    let page = cached?.page;
    let nodes = cached?.nodes;
    let variants = cached?.variants;

    if (!cached) {
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

      // Components drawn against a theme are previewed in it; the rest get the
      // default, which is what a project has before it chooses one.
      const style = themeById(component.themeHint ?? "clean-minimal").style;

      const render = async (data) => {
        // Placeholder mode, so a generated node shows what an author sees rather
        // than requiring an endpoint to build the docs.
        const document = await buildDocument(
          wrap(component.id, component.title, exported),
          data,
          { dynamicMode: "placeholder" },
        );

        return {
          nodes: document.nodes,
          page: await renderDocxPage({ ...document, style }, {
            title: component.title,
            style: NODE_ONLY_STYLE,
          }),
        };
      };

      ({ page, nodes } = await render(component.previewData));
      variants = await renderVariants(component, render, { page, nodes });

      await writeCache(key, { page, nodes, variants });
    } else {
      reused += 1;
    }

    await writeFile(resolve(COMPONENT_OUT, `${component.id}.html`), page, "utf8");
    for (const variant of variants?.values ?? []) {
      await writeFile(resolve(COMPONENT_OUT, variantFile(component.id, variant.value)), variant.page, "utf8");
    }

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
      resolved: nodes,
      previewUrl: posix.join(COMPONENT_BASE, `${component.id}.html`),
      // Only where each page was written goes in the manifest, not the page.
      variants: variants && {
        path: variants.path,
        selected: variants.selected,
        values: variants.values.map(({ page: _page, ...variant }) => ({
          ...variant,
          previewUrl: posix.join(COMPONENT_BASE, variantFile(component.id, variant.value)),
        })),
      },
    });

    console.log(
      `registry: component ${component.id.padEnd(16)} ${nodes.length} nodes` +
        `${variants ? `, ${variants.values.length} ${variants.path} previews` : ""}` +
        `${cached ? " (cached)" : ""}`,
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
    `registry: ${themes.length} themes, ${components.length} components ` +
      `(${reused} reused from cache) -> public/demo/, ` +
      `manifest -> src/generated/registry.json`,
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
 * The field a component's previews are worth flipping through, or null.
 *
 * The first data field typed as a union of two or more string literals — a
 * treatment, a kind, a tone — whose value is set in the preview data. A
 * component branches on a field like that, so one preview of one value hides
 * most of what it can print. Read from the catalog's type rather than parsed
 * out of the node's source: it is the same union the page already shows the
 * reader, and it is a string here rather than a TypeScript AST.
 */
function variantAxis(component) {
  for (const field of component.dataFields) {
    const literals = field.type
      .split(" | ")
      .filter((member) => /^"[^"]*"$/.test(member))
      .map((member) => member.slice(1, -1));
    const current = readPath(component.previewData, field.path);

    if (literals.length >= 2 && literals.includes(current)) {
      return { path: field.path, values: literals, selected: current };
    }
  }

  return null;
}

/**
 * The component once per value of its axis, the rest of the preview data
 * unchanged. The value the preview data already holds reuses the render made
 * for it rather than drawing it twice.
 */
async function renderVariants(component, render, base) {
  const axis = variantAxis(component);
  if (!axis) return null;

  const values = [];

  for (const value of axis.values) {
    const previewData = structuredClone(component.previewData);
    writePath(previewData, axis.path, value);

    const { page, nodes } = value === axis.selected ? base : await render(previewData);
    values.push({ value, previewData, resolved: nodes, page });
  }

  return { path: axis.path, selected: axis.selected, values };
}

/** The values are slugs today, but a file name is no place to rely on it. */
function variantFile(id, value) {
  return `${id}--${value.replace(/[^a-z0-9-]/gi, "_")}.html`;
}

function readPath(data, path) {
  return path.split(".").reduce((value, key) => value?.[key], data);
}

function writePath(data, path, value) {
  const keys = path.split(".");
  const parent = keys.slice(0, -1).reduce((object, key) => object[key], data);
  parent[keys.at(-1)] = value;
}

/**
 * What a preview was built from, as one string.
 *
 * Everything that could change the output goes in: the sources, the data they
 * were resolved against, the framework that resolved them, and the version of
 * this script. Anything left out is a way for the cache to serve something that
 * no longer matches what the registry holds.
 */
function cacheKey(parts) {
  return createHash("sha256")
    .update(JSON.stringify({ version: CACHE_VERSION, ...parts }))
    .digest("hex")
    .slice(0, 16);
}

/** A cached build, or null when this is the first time these inputs were seen. */
async function readCache(key) {
  try {
    return JSON.parse(await readFile(resolve(CACHE_DIR, `${key}.json`), "utf8"));
  } catch {
    return null;
  }
}

async function writeCache(key, value) {
  await writeFile(resolve(CACHE_DIR, `${key}.json`), JSON.stringify(value), "utf8");
}

/**
 * The installed framework's version, so upgrading it invalidates every preview.
 *
 * A renderer that starts drawing tables differently has to redraw the previews,
 * and nothing in a component's own source would say so.
 */
async function frameworkVersion() {
  // Read from the package directory rather than resolved through the exports
  // map, which does not expose package.json — and the site depends on the
  // package as `file:..`, so this is the one being previewed.
  return JSON.parse(await readFile(resolve(ROOT, "..", "package.json"), "utf8")).version;
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
