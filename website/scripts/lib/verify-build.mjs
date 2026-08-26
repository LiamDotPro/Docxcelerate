/**
 * The build step of the invoice verification harness.
 *
 * Produces the artefacts every probe measures: the packed `.docx` (what Word
 * opens), the baked preview page (what docx-preview lays out), and — behind
 * flags — the variant builds specific objectives compare against. One build
 * feeding every probe is the point: a probe that built its own document could
 * disagree with the others about what was measured.
 *
 * See VERIFY-CONTRACT.md ("The build step") for the flags and their consumers.
 */
import { build } from "esbuild";
import { docxcelerateEsbuildTransform } from "docxcelerate/transform";
import { mkdir, rm, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { PAGE_ONLY_STYLE, renderDocxPage } from "./docx-page.mjs";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..", "..");
const VERIFY_DIR = resolve(ROOT, ".verify");
const PROJECT_SOURCE = resolve(ROOT, "src/demo/documents/invoice/document.project.ts");

/** Bundles and imports the invoice project, exactly as the demo build does. */
async function loadProject() {
  const tempEntry = resolve(VERIFY_DIR, "project-bundle.mjs");

  await build({
    entryPoints: [PROJECT_SOURCE],
    outfile: tempEntry,
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

  // A cache-busting query, so repeated runs in one process see fresh builds.
  const mod = await import(`${pathToFileURL(tempEntry).href}?t=${Date.now()}`);

  return mod.default;
}

/**
 * The document's derivers, bundled so Node can import them.
 *
 * A published document names its derivers and does not carry them, so
 * resolving one means holding the registry an engine would be handed. They
 * live in TypeScript beside the document, hence the same esbuild pass the
 * project itself goes through.
 */
async function loadDerivers() {
  const entry = resolve(VERIFY_DIR, "derivers-bundle.mjs");

  await build({
    entryPoints: [resolve(ROOT, "src/demo/documents/invoice/derivers.ts")],
    outfile: entry,
    bundle: true,
    format: "esm",
    platform: "node",
    target: "node20",
    packages: "external",
    logLevel: "warning",
  });

  return await import(`${pathToFileURL(entry).href}?t=${Date.now()}`);
}

/** The preview data, optionally reshaped for a variant build. */
function variantData(project, { lines } = {}) {
  const data = structuredClone(project.previewData);

  if (typeof lines === "number") {
    const base = data.lines;
    data.lines = Array.from({ length: lines }, (_, index) => {
      const line = base[index % base.length];
      return index < base.length ? line : {
        ...line,
        desc: `${line.desc} (${Math.floor(index / base.length) + 1})`,
      };
    });
  }

  return data;
}

/**
 * Builds the invoice and writes the probe inputs.
 *
 * @param {{ lines?: number, publish?: boolean }} options Variant flags.
 * @returns What was written, plus the built model for callers that verify it.
 */
export async function buildVerificationArtifacts(options = {}) {
  await mkdir(VERIFY_DIR, { recursive: true });

  const project = await loadProject();
  const docxcelerate = await import("docxcelerate");
  const { createDocxBlob } = await import("docxcelerate/docx");

  const suffix = typeof options.lines === "number" ? `-lines${options.lines}` : "";
  const previewData = variantData(project, options);
  const model = await docxcelerate.buildProjectPreviewDocument(
    { ...project, previewData },
  );

  const blob = await createDocxBlob(model);
  const docxPath = resolve(VERIFY_DIR, `invoice${suffix}.docx`);
  await writeFile(docxPath, new Uint8Array(await blob.arrayBuffer()));

  const html = await renderDocxPage(model, {
    title: project.name,
    style: PAGE_ONLY_STYLE,
  });
  const htmlPath = resolve(VERIFY_DIR, `invoice${suffix}.html`);
  await writeFile(htmlPath, html, "utf8");

  const modelPath = resolve(VERIFY_DIR, `model${suffix}.json`);
  await writeFile(modelPath, JSON.stringify(model, null, 2), "utf8");

  const result = { docxPath, htmlPath, modelPath, model };

  if (options.publish) {
    // The publish-path build. Until D13 lands this throws (`reduce` is refused
    // by the publish stand-in); the runner records that as BLOCKED evidence
    // rather than a harness failure.
    const publishDir = resolve(VERIFY_DIR, "publish");
    await rm(publishDir, { recursive: true, force: true });
    await mkdir(publishDir, { recursive: true });

    try {
      const engine = await docxcelerate.buildProjectEngineDocument(project);
      await writeFile(
        resolve(publishDir, "document.json"),
        JSON.stringify(engine, null, 2),
        "utf8",
      );
      result.publish = { ok: true, documentPath: resolve(publishDir, "document.json") };

      // The round trip the publish path actually promises: take the published
      // document, resolve it against the same data an engine would be handed,
      // and pack that. A striping decision baked at build time survives the
      // first build and dies here, which is the whole point of measuring it.
      // The registry an engine would be handed: the document's own derivers,
      // named as the published tokens name them.
      const registry = docxcelerate.createDefaultDeriverRegistry();
      for (const entry of Object.values(await loadDerivers())) {
        if (entry && typeof entry === "object" && typeof entry.run === "function") {
          registry.register(entry.name, entry.run, entry.placeholder);
        }
      }

      const settled = await docxcelerate.resolveDocument(
        engine,
        {
          ctx: {},
          derived: {},
          dataProvider: new docxcelerate.InMemoryDataProvider(previewData),
          aiClient: {
            generateParagraph: ({ node }) => `[dynamic: ${node.id}]`,
            generateImage: () => { throw new Error("the published invoice should ask for no image"); },
          },
        },
        { derivers: registry },
      );
      const publishedBlob = await createDocxBlob({ ...settled, style: model.style });
      const publishedDocx = resolve(publishDir, "invoice-published.docx");
      await writeFile(publishedDocx, new Uint8Array(await publishedBlob.arrayBuffer()));
      result.publish.docxPath = publishedDocx;
    } catch (error) {
      result.publish = { ok: false, error: String(error?.message ?? error) };
      await writeFile(
        resolve(publishDir, "error.txt"),
        result.publish.error,
        "utf8",
      );
    }
  }

  return result;
}

/**
 * F9's minimal fixture: does a block that names no edges draw no edges?
 *
 * The invoice cannot answer this on its own — a stray rule there could come
 * from the theme, the row separator or the block, and the objective is about
 * the block alone. So a two-cell document is packed with one variant that
 * asks for a border on no sides at all, and the packed XML is asked whether a
 * `w:tcBorders` was written anyway. Nothing about the invoice is involved.
 */
export async function buildBorderFixture() {
  const { createDocxBlob } = await import("docxcelerate/docx");
  const { cleanMinimalDocumentStyle } = await import("docxcelerate");

  const style = {
    ...cleanMinimalDocumentStyle,
    blocks: {
      // A border colour with an empty side list: the colour is set, and no
      // edge is asked for. Nothing should be drawn.
      "no-edges": { border: "D9DDEB", borderSides: [], paddingPt: 12 },
    },
  };

  const model = {
    id: "border-fixture",
    title: "Border fixture",
    style,
    nodes: [
      {
        kind: "table",
        id: "t",
        columns: [{ width: "auto" }],
        children: [
          {
            kind: "tableRow",
            id: "r",
            children: [
              { kind: "tableCell", id: "c", variant: "no-edges", children: [
                { kind: "paragraph", id: "p", text: "No edges." },
              ] },
            ],
          },
        ],
      },
    ],
  };

  const blob = await createDocxBlob(model);
  const bytes = Buffer.from(await blob.arrayBuffer());
  const path = resolve(VERIFY_DIR, "border-fixture.docx");
  await mkdir(VERIFY_DIR, { recursive: true });
  await writeFile(path, new Uint8Array(bytes));

  // Read word/document.xml straight back out of the package rather than
  // trusting the model: the objective is about what was packed.
  const { entryOf } = await import("./probe-ooxml.mjs");
  const documentXml = new TextDecoder().decode(entryOf(bytes, "word/document.xml"));

  return {
    path,
    emitsTcBorders: /<w:tcBorders[ />]/.test(documentXml),
  };
}

// CLI: node scripts/lib/verify-build.mjs [--lines=N] [--publish]
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const options = {};
  for (const arg of process.argv.slice(2)) {
    if (arg.startsWith("--lines=")) options.lines = Number(arg.slice(8));
    if (arg === "--publish") options.publish = true;
  }
  buildVerificationArtifacts(options).then((result) => {
    console.log("built:", result.docxPath);
    console.log("built:", result.htmlPath);
    if (result.publish) console.log("publish:", JSON.stringify(result.publish));
  }, (error) => {
    console.error("verify-build failed:", error);
    process.exitCode = 1;
  });
}
