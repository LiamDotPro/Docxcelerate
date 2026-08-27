/**
 * Turning a `.case.tsx` into the artefacts every probe reads.
 *
 * One build feeds all four probes. A probe that built its own copy could
 * disagree with the others about what was measured, and a board built on that
 * disagreement is worse than no board.
 *
 * @module
 */

import { build } from "esbuild";
import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url));
export const ROOT = resolve(HERE, "..");
export const CASES_DIR = resolve(ROOT, "cases");
export const OUT_DIR = resolve(ROOT, ".out");

/**
 * Bundles one case file and imports it.
 *
 * The same esbuild recipe the site's demo build uses: the framework's own
 * transform first (so a case may use hooks like any document), then automatic
 * JSX pointed at the template runtime, with the package left external so the
 * case runs against `dist/` rather than a second copy of it.
 */
export async function loadCase(sourcePath) {
  const { docxcelerateEsbuildTransform } = await import("docxcelerate/transform");
  const bundle = resolve(OUT_DIR, "_bundles", `${hashOf(sourcePath)}.mjs`);

  await mkdir(dirname(bundle), { recursive: true });
  await build({
    entryPoints: [sourcePath],
    outfile: bundle,
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

  // A cache-busting query, so one process can rebuild a case it already ran.
  const module = await import(`${pathToFileURL(bundle).href}?t=${Date.now()}`);
  const spec = module.default;

  if (spec === undefined) {
    throw new Error(`${sourcePath}: no default export — a case file exports defineCase(...)`);
  }

  return spec;
}

/**
 * Builds one case: the model, the `.docx` Word opens, and the HTML the preview
 * lays out.
 *
 * @param {object} spec A case, as `defineCase` returned it.
 * @returns Paths and the built model.
 */
export async function buildCase(spec) {
  const { buildDocument } = await import("docxcelerate");
  const { createDocxBlob } = await import("docxcelerate/docx");
  const { renderPreviewPage } = await import("./preview-render.mjs");

  const outDir = resolve(OUT_DIR, spec.id);
  await mkdir(outDir, { recursive: true });

  const built = await buildDocument(spec.document, spec.data ?? {}, {
    branchMode: "decide",
    dynamicMode: "placeholder",
  });
  const model = spec.style === undefined ? built : { ...built, style: spec.style };

  const blob = await createDocxBlob(model);
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const docxPath = resolve(outDir, "case.docx");
  await writeLocked(docxPath, bytes);

  const htmlPath = resolve(outDir, "case.html");
  await writeFile(htmlPath, await renderPreviewPage(model, spec.title), "utf8");

  const modelPath = resolve(outDir, "model.json");
  await writeFile(modelPath, JSON.stringify(model, null, 2), "utf8");

  return { outDir, docxPath, htmlPath, modelPath, model, bytes };
}

/**
 * Writes a file Word may still be holding.
 *
 * Word releases a document a moment after it is told to close it, not at the
 * instant — and a run that measures a case and then rebuilds it can arrive
 * inside that moment, where Windows answers `EBUSY` and the case fails for a
 * reason that has nothing to do with the document. Seen once on a full board,
 * which is once more than a harness should ever fail for its own reasons.
 *
 * Three tries over a second and a half. Anything longer is a lock somebody
 * should know about rather than one to wait out, so it is allowed to throw.
 */
async function writeLocked(path, bytes) {
  for (let attempt = 1; ; attempt += 1) {
    try {
      await writeFile(path, bytes);
      return;
    } catch (error) {
      if (attempt >= 3 || (error?.code !== "EBUSY" && error?.code !== "EPERM")) {
        throw error;
      }
      await new Promise((done) => setTimeout(done, attempt * 500));
    }
  }
}

/** A short, stable name for a bundle, so two cases never collide on disk. */
function hashOf(value) {
  let hash = 5381;
  for (let index = 0; index < value.length; index += 1) {
    hash = ((hash * 33) ^ value.charCodeAt(index)) >>> 0;
  }
  return hash.toString(36);
}
