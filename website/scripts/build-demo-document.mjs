/**
 * Bakes the homepage hero preview.
 *
 * The hero shows a real document project beside the document it produces. Both
 * halves come from ONE source tree — src/demo/documents/<id>/. This script
 * executes it; the page imports the same files with ?raw to display them in
 * tabs. They cannot drift, and if the renderer changes, the site changes.
 *
 * Output: public/demo/<id>.html — the packed .docx, read back and laid out by
 * docx-preview, so the pane shows the file itself rather than a second
 * renderer's idea of it. Embedded in an iframe, so its styles stay isolated
 * from the site's.
 */
import { build } from "esbuild";
import { docxcelerateEsbuildTransform } from "docxcelerate/transform";
import { existsSync } from "node:fs";
import { mkdir, readdir, rm, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { PAGE_ONLY_STYLE, renderDocxPage } from "./lib/docx-page.mjs";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..");

const DOCUMENTS_DIR = resolve(ROOT, "src/demo/documents");
const TEMP_DIR = resolve(ROOT, ".demo-build");
const OUT_DIR = resolve(ROOT, "public/demo");

async function main() {
  await rm(TEMP_DIR, { recursive: true, force: true });
  await mkdir(TEMP_DIR, { recursive: true });
  await mkdir(OUT_DIR, { recursive: true });

  const entries = await readdir(DOCUMENTS_DIR, { withFileTypes: true });
  // `npm run demo -- portfolio-review` builds one document; no argument builds them all.
  // `--docx <dir>` also writes each packed file there, for checking in Word.
  const args = process.argv.slice(2);
  const docxAt = args.includes("--docx") ? resolve(args[args.indexOf("--docx") + 1]) : null;
  const only = args.filter((arg, i) => arg !== "--docx" && args[i - 1] !== "--docx");
  const dirs = entries
    .filter((entry) => entry.isDirectory() && (only.length === 0 || only.includes(entry.name)))
    .map((entry) => entry.name)
    .sort();

  if (dirs.length === 0) {
    throw new Error(`No document projects found in ${DOCUMENTS_DIR}`);
  }

  const { buildProjectFinalDocument, buildProjectPreviewDocument } = await import("docxcelerate");
  const { createDocxBlob } = await import("docxcelerate/docx");
  const writeDocx = async (name, model) => {
    if (docxAt === null) return;
    await mkdir(docxAt, { recursive: true });
    await writeFile(resolve(docxAt, `${name}.docx`), new Uint8Array(await (await createDocxBlob(model)).arrayBuffer()));
  };

  for (const dir of dirs) {
    const source = resolve(DOCUMENTS_DIR, dir, "document.project.ts");
    const tempEntry = resolve(TEMP_DIR, `${dir}.mjs`);
    // A document that ships the engine's answers for its preview client gets a
    // second page: the same file as the engine would write it for that client.
    const engineSource = resolve(DOCUMENTS_DIR, dir, "engine-text.ts");
    const engineEntry = resolve(TEMP_DIR, `${dir}.engine.mjs`);

    // Bundle so the project's relative .ts imports resolve, but keep
    // `docxcelerate` external so it loads through website/node_modules -> the
    // parent package's exports map, exactly as a real consumer's would.
    await build({
      entryPoints: existsSync(engineSource)
        ? { [dir]: source, [`${dir}.engine`]: engineSource }
        : { [dir]: source },
      outdir: TEMP_DIR,
      outExtension: { ".js": ".mjs" },
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

    const mod = await import(pathToFileURL(tempEntry).href);
    const project = mod.default;

    if (project.id !== dir) {
      throw new Error(
        `Document id "${project.id}" does not match its directory "${dir}". ` +
          `The site addresses previews by directory name.`,
      );
    }

    const document = await buildProjectPreviewDocument(project);
    const html = await renderDocxPage(document, {
      title: project.name,
      style: PAGE_ONLY_STYLE,
    });

    await writeFile(resolve(OUT_DIR, `${dir}.html`), html, "utf8");
    await writeDocx(dir, document);

    if (existsSync(engineSource)) {
      const { engineText } = await import(pathToFileURL(engineEntry).href);
      const written = await buildProjectFinalDocument(project, {
        aiClient: {
          generateParagraph: ({ node }) => {
            // A node a loop wrote is numbered for its entry; the answer is
            // written once for the node, so the number comes off to find it.
            const answer = engineText[node.id] ?? engineText[node.id.replace(/-\d+$/, "")];
            if (answer === undefined) throw new Error(`engine-text.ts has no answer for "${node.id}"`);
            return answer;
          },
        },
      });
      const engineHtml = await renderDocxPage(written, { title: project.name, style: PAGE_ONLY_STYLE });
      await writeFile(resolve(OUT_DIR, `${dir}.engine.html`), engineHtml, "utf8");
      await writeDocx(`${dir}.engine`, written);
      console.log(`demo: ${project.name.padEnd(20)} engine build -> public/demo/${dir}.engine.html`);
    }
    console.log(
      `demo: ${project.name.padEnd(20)} ${String(countNodes(document.nodes)).padStart(2)} nodes` +
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
