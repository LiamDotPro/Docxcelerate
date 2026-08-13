/**
 * Bakes the homepage hero preview.
 *
 * The hero shows a real letter project beside the letter it produces. Both
 * halves come from ONE source tree — src/demo/offer-of-admission/. This script
 * executes it; the page imports the same files with ?raw to display them in
 * tabs. They cannot drift, and if the renderer changes, the site changes.
 *
 * Output: public/demo/letter.html — the renderer's own standalone document,
 * with the Word-style app chrome hidden so the pane shows only the page.
 * Embedded in an iframe, so its styles stay isolated from the site's.
 */
import { build } from "esbuild";
import { mkdir, readdir, rm, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { extractPage, PAGE_ONLY_STYLE } from "./lib/preview-page.mjs";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..");

const LETTERS_DIR = resolve(ROOT, "src/demo/letters");
const TEMP_DIR = resolve(ROOT, ".demo-build");
const OUT_DIR = resolve(ROOT, "public/demo");

async function main() {
  await rm(TEMP_DIR, { recursive: true, force: true });
  await mkdir(TEMP_DIR, { recursive: true });
  await mkdir(OUT_DIR, { recursive: true });

  const entries = await readdir(LETTERS_DIR, { withFileTypes: true });
  const dirs = entries.filter((entry) => entry.isDirectory()).map((entry) => entry.name).sort();

  if (dirs.length === 0) {
    throw new Error(`No letter projects found in ${LETTERS_DIR}`);
  }

  const [{ buildProjectPreviewDocument }, { renderDocumentWebsite }] = await Promise.all([
    import("docxcelerate"),
    import("docxcelerate/renderer"),
  ]);

  for (const dir of dirs) {
    const source = resolve(LETTERS_DIR, dir, "letter.project.ts");
    const tempEntry = resolve(TEMP_DIR, `${dir}.mjs`);

    // Bundle so the project's relative .ts imports resolve, but keep
    // `docxcelerate` external so it loads through website/node_modules -> the
    // parent package's exports map, exactly as a real consumer's would.
    await build({
      entryPoints: [source],
      outfile: tempEntry,
      bundle: true,
      format: "esm",
      platform: "node",
      target: "node20",
      jsx: "automatic",
      jsxImportSource: "docxcelerate/template",
      packages: "external",
      logLevel: "warning",
    });

    const mod = await import(pathToFileURL(tempEntry).href);
    const project = mod.default;

    if (project.id !== dir) {
      throw new Error(
        `Letter id "${project.id}" does not match its directory "${dir}". ` +
          `The site addresses previews by directory name.`,
      );
    }

    const letter = await buildProjectPreviewDocument(project);
    const rendered = renderDocumentWebsite(letter, { title: project.name });

    const html = extractPage(rendered, { title: project.name, style: PAGE_ONLY_STYLE });

    await writeFile(resolve(OUT_DIR, `${dir}.html`), html, "utf8");
    console.log(
      `demo: ${project.name.padEnd(20)} ${String(countNodes(letter.nodes)).padStart(2)} nodes` +
        ` -> public/demo/${dir}.html`,
    );
  }

  await rm(TEMP_DIR, { recursive: true, force: true });
}

function countNodes(nodes) {
  return nodes.reduce(
    (total, node) => total + 1 + (node.children ? countNodes(node.children) : 0),
    0,
  );
}

main().catch((error) => {
  console.error("demo build failed:", error);
  process.exitCode = 1;
});
