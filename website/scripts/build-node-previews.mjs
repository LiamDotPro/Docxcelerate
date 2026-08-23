/**
 * Bakes the node gallery.
 *
 * src/nodes/ is a real letter project's worth of node components, one file per
 * variant, described by src/nodes/catalog.ts. This script resolves every
 * variant through the framework itself and emits two things:
 *
 *   public/demo/nodes/<type>/<variant>.html   the node as the packed .docx
 *                                             renders it
 *   src/generated/node-catalog.json           what each variant resolved to,
 *                                             plus the catalog's own metadata
 *
 * The docs pages and the homepage read only the JSON. If a helper changes
 * shape, this build fails or the previews change — either way the site follows.
 */
import { build } from "esbuild";
import { access, mkdir, rm, writeFile } from "node:fs/promises";
import { dirname, posix, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { NODE_ONLY_STYLE, renderDocxPage } from "./lib/docx-page.mjs";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..");

const NODES_DIR = resolve(ROOT, "src/nodes");
const TEMP_DIR = resolve(ROOT, ".demo-build");
const OUT_DIR = resolve(ROOT, "public/demo/nodes");
const MANIFEST = resolve(ROOT, "src/generated/node-catalog.json");

/** Where previews are addressed from in the browser. */
const PREVIEW_BASE = "/demo/nodes";

async function main() {
  await rm(TEMP_DIR, { recursive: true, force: true });
  // Wiped rather than overwritten: a deleted variant should take its preview
  // with it instead of leaving a page nothing links to.
  await rm(OUT_DIR, { recursive: true, force: true });
  await mkdir(TEMP_DIR, { recursive: true });
  await mkdir(OUT_DIR, { recursive: true });
  await mkdir(dirname(MANIFEST), { recursive: true });

  const { NODE_CATEGORIES, NODE_TYPES, sampleData } = await loadCatalog();
  const { buildDocument } = await import("docxcelerate");
  const one = await singleNodeTemplate();

  const types = [];
  let previews = 0;

  for (const type of NODE_TYPES) {
    const variants = [];

    for (const variant of type.variants) {
      const sourceFile = `${type.id}/${variant.id}.node.tsx`;
      await assertSourceExists(sourceFile, type.id, variant.id);

      // One node, resolved exactly as a document would resolve it. Placeholder
      // mode is what preview builds use, so a dynamic node shows what an
      // author sees rather than requiring an engine to render docs.
      const document = await buildDocument(
        one(`${type.id}-${variant.id}`, variant.title, variant.component),
        sampleData,
        { dynamicMode: "placeholder" },
      );

      const out = resolve(OUT_DIR, type.id, `${variant.id}.html`);
      await mkdir(dirname(out), { recursive: true });
      await writeFile(
        out,
        await renderDocxPage(document, { title: variant.title, style: NODE_ONLY_STYLE }),
        "utf8",
      );
      previews += 1;

      variants.push({
        id: variant.id,
        title: variant.title,
        summary: variant.summary,
        sourceFile,
        previewUrl: posix.join(PREVIEW_BASE, type.id, `${variant.id}.html`),
        // What the helper actually produced. Shown on the node's page, so the
        // shape a renderer or endpoint receives is visible rather than
        // described.
        resolved: document.nodes[0],
        prompts: await resolvePrompts(variant, { one, buildDocument, sampleData }),
      });
    }

    types.push({
      id: type.id,
      title: type.title,
      kind: type.kind,
      category: type.category,
      status: type.status,
      helpers: type.helpers,
      summary: type.summary,
      detail: type.detail,
      children: type.children,
      resolves: type.resolves,
      renderNote: type.renderNote ?? null,
      options: type.options,
      variants,
    });

    console.log(
      `nodes: ${type.title.padEnd(20)} ${String(variants.length).padStart(2)} variants` +
        (type.status === "planned" ? "  (planned)" : ""),
    );
  }

  await writeFile(
    MANIFEST,
    `${JSON.stringify({ categories: NODE_CATEGORIES, types }, null, 2)}\n`,
    "utf8",
  );

  console.log(
    `nodes: ${types.length} types, ${previews} previews -> public${PREVIEW_BASE}/, ` +
      `manifest -> src/generated/node-catalog.json`,
  );

  await rm(TEMP_DIR, { recursive: true, force: true });
}

/**
 * The other half of a dynamic node: the prompts, resolved against the same
 * sample data, as a generation endpoint would be handed them.
 *
 * Preview builds never produce these — they stop at the placeholder — so the
 * node is resolved a second time with a client that returns nothing. What
 * comes back is the request minus the response, which is exactly the part
 * worth reading on a docs page. Static nodes have none, and neither do
 * sections; both return null and the page shows no prompt panel.
 */
async function resolvePrompts(variant, { one, buildDocument, sampleData }) {
  const document = await buildDocument(
    one("prompts", variant.title, variant.component),
    sampleData,
    { dynamicMode: "resolve", aiClient: EMPTY_AI_CLIENT },
  );

  const node = document.nodes[0];

  return node?.prompts?.length ? node.prompts : null;
}

/**
 * Wraps one component in a document, without JSX.
 *
 * This script is plain JavaScript, so it reaches for the same jsx() call a
 * compiled `<Document>` would make. Building the element by hand rather than
 * adding a build step keeps the gallery resolving through the real framework,
 * which is the whole point of generating it.
 */
async function singleNodeTemplate() {
  const { jsx } = await import("docxcelerate/template/jsx-runtime");
  const { Document, template } = await import("docxcelerate/template");

  return (id, title, component) =>
    template(jsx(Document, { id, title, children: jsx(component, {}) }));
}

/** Answers nothing, so only the prompts survive the round trip. */
const EMPTY_AI_CLIENT = {
  generateParagraph: () => "",
  generateImage: () => ({ path: "" }),
  generateGraph: () => ({}),
};

/**
 * Bundles the catalog so its relative .ts imports resolve, keeping
 * `docxcelerate` external so it loads through website/node_modules -> the
 * parent package's exports map, exactly as a real consumer's would.
 */
async function loadCatalog() {
  const entry = resolve(TEMP_DIR, "node-catalog.mjs");

  await build({
    stdin: {
      contents:
        `export { NODE_CATEGORIES, NODE_TYPES } from "./catalog.ts";\n` +
        `export { sampleData } from "./sample-data.ts";\n`,
      resolveDir: NODES_DIR,
      sourcefile: "node-catalog-entry.ts",
      loader: "ts",
    },
    outfile: entry,
    bundle: true,
    format: "esm",
    platform: "node",
    target: "node20",
    packages: "external",
    logLevel: "warning",
  });

  return await import(pathToFileURL(entry).href);
}

/**
 * Variant ids address a source file and a preview by convention. Checking the
 * file exists turns a rename that missed one half into a build failure rather
 * than a docs page quietly showing the wrong code.
 */
async function assertSourceExists(sourceFile, typeId, variantId) {
  try {
    await access(resolve(NODES_DIR, sourceFile));
  } catch {
    throw new Error(
      `Variant "${variantId}" of node type "${typeId}" has no source at ` +
        `src/nodes/${sourceFile}. Variant ids name their own file.`,
    );
  }
}

main().catch((error) => {
  console.error("node preview build failed:", error);
  process.exitCode = 1;
});
