/**
 * Word's PDF export, turned into the page pictures the board shows.
 *
 * Node has no canvas, so the rasterising happens where one exists: a harness
 * page served to headless Chrome loads pdf.js from `node_modules`, draws each
 * page at 96dpi — the same scale the preview screenshots are taken at, so the
 * two columns of the board are directly comparable — and `--dump-dom` carries
 * the pixels back as data URIs for Node to decode.
 *
 * @module
 */

import { execFile } from "node:child_process";
import { createServer } from "node:http";
import { copyFile, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, extname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { setTimeout as sleep } from "node:timers/promises";
import { chromePath } from "./probe-preview.mjs";

const execFileAsync = promisify(execFile);
const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..");
const PORT = 8903;

/**
 * How many pages of a document the board shows.
 *
 * Each page comes back through `--dump-dom` as roughly a megabyte of base64,
 * and a five-page document never finished. Three sheets is what a conformance
 * case needs to show; the run reports how many it left behind, because a cap
 * nobody is told about reads as "that was all there was".
 */
const PAGE_CAP = 3;

/** Where pdf.js lives, if it was installed. */
function pdfjsPaths() {
  const base = resolve(ROOT, "node_modules", "pdfjs-dist", "build");
  const library = resolve(base, "pdf.mjs");
  const worker = resolve(base, "pdf.worker.mjs");

  return existsSync(library) && existsSync(worker) ? { library, worker } : null;
}

/**
 * The rasterising page: pdf.js draws one page onto one canvas, and Chrome
 * screenshots it.
 *
 * The pixels do not come back through `--dump-dom`. They used to, as a base64
 * data URI per page, and it worked for one page and for two and then stopped:
 * at three pages the dumped DOM never arrived, with no error to say why. So
 * each page is now rendered on its own, alone on the page, at exactly its own
 * size — and `--screenshot` writes the PNG straight to disk, which has no such
 * ceiling. One Chrome run per page, plus a cheap first run that asks only how
 * many pages there are and how big they are.
 *
 * @param page Which page to draw, or `0` to report the document's shape and
 * draw nothing.
 */
function harnessHtml(page) {
  return `<!doctype html>
<html>
<head><meta charset="utf-8"><title>raster</title></head>
<body style="margin:0">
<div id="pages"></div>
<script>
// A module that fails to load runs none of its own error handling, so the
// window is what reports it — otherwise the harness comes back with an empty
// DOM and no reason for it.
window.addEventListener("error", function (event) {
  var pre = document.createElement("pre");
  pre.id = "raster-error";
  pre.textContent = "window: " + String(event.message || event.error || event);
  document.body.appendChild(pre);
});
window.addEventListener("unhandledrejection", function (event) {
  var pre = document.createElement("pre");
  pre.id = "raster-error";
  pre.textContent = "rejected: " + String((event.reason && event.reason.stack) || event.reason);
  document.body.appendChild(pre);
});
</script>
<script type="module">
const WANTED_PAGE = ${page};
import * as pdfjs from "./pdf.mjs";
// The worker module is imported and handed over as window.pdfjsWorker, which
// is pdf.js's own signal to run it on this thread rather than spawning a real
// Worker. It has to be this way round: Chrome's --virtual-time-budget
// fast-forwards the page's clock but not a worker thread's, so a real worker
// never gets to reply and the render sits pending until the budget expires --
// which looks exactly like a page that produced nothing, with no error to say
// why. On the main thread the whole render happens inside virtual time.
import * as pdfjsWorker from "./pdf.worker.mjs";
window.pdfjsWorker = pdfjsWorker;

(async function () {
  try {
    const doc = await pdfjs.getDocument({ url: "./case.pdf" }).promise;
    const host = document.getElementById("pages");

    // 96/72: the scale that puts a point-measured PDF on the same pixel grid
    // the preview screenshots are taken at, so the board's two columns are
    // directly comparable.
    const first = await doc.getPage(1);
    const shape = first.getViewport({ scale: 96 / 72 });

    if (WANTED_PAGE > 0 && WANTED_PAGE <= doc.numPages) {
      const page = await doc.getPage(WANTED_PAGE);
      const viewport = page.getViewport({ scale: 96 / 72 });
      const canvas = document.createElement("canvas");
      canvas.width = Math.round(viewport.width);
      canvas.height = Math.round(viewport.height);
      canvas.style.display = "block";
      host.appendChild(canvas);
      await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
    }

    const done = document.createElement("div");
    done.id = "raster-done";
    done.textContent = doc.numPages + "x" + Math.round(shape.width) + "x" + Math.round(shape.height);
    document.body.appendChild(done);
  } catch (error) {
    const pre = document.createElement("pre");
    pre.id = "raster-error";
    pre.textContent = String((error && error.stack) || error);
    document.body.appendChild(pre);
  }
})();
</script>
</body>
</html>
`;
}

async function startServer(root) {
  const types = {
    ".html": "text/html; charset=utf-8",
    ".mjs": "text/javascript; charset=utf-8",
    ".pdf": "application/pdf",
  };

  for (let attempt = 1; ; attempt += 1) {
    try {
      return await new Promise((done, fail) => {
        const server = createServer(async (request, response) => {
          try {
            const pathname = decodeURIComponent(new URL(request.url, "http://localhost").pathname);
            const file = join(root, pathname.replaceAll("..", ""));
            const body = await readFile(file);
            response.writeHead(200, {
              "Content-Type": types[extname(file)] ?? "application/octet-stream",
            });
            response.end(body);
          } catch {
            response.writeHead(404);
            response.end("not found");
          }
        });
        server.once("error", fail);
        server.listen(PORT, "127.0.0.1", () => done(server));
      });
    } catch (error) {
      if (error?.code !== "EADDRINUSE" || attempt >= 4) throw error;
      await sleep(1000);
    }
  }
}

/**
 * Rasters `case.pdf` into `word-p1.png`, `word-p2.png`, … beside it.
 *
 * @returns What was written. An empty list when there is nothing to raster or
 * pdf.js is not installed — a missing picture costs the board a column, not a
 * run.
 */
export async function rasterisePdf(outDir) {
  const pdfPath = resolve(outDir, "case.pdf");
  const chrome = chromePath();
  const pdfjs = pdfjsPaths();

  if (!existsSync(pdfPath) || chrome === null || pdfjs === null) {
    return [];
  }

  let server;
  let tempDir;

  try {
    tempDir = await mkdtemp(join(tmpdir(), "conf-raster-"));
    await copyFile(pdfPath, join(tempDir, "case.pdf"));
    await copyFile(pdfjs.library, join(tempDir, "pdf.mjs"));
    await copyFile(pdfjs.worker, join(tempDir, "pdf.worker.mjs"));

    server = await startServer(tempDir);

    /** One Chrome run against the harness, asking it to draw one page. */
    const draw = async (page, extra = []) => {
      await writeFile(join(tempDir, "index.html"), harnessHtml(page), "utf8");

      const { stdout } = await execFileAsync(chrome, [
        "--headless=new",
        "--disable-gpu",
        "--hide-scrollbars",
        "--force-device-scale-factor=1",
        // Generous, because the budget is virtual: Chrome fast-forwards the
        // page's clock, so a larger number costs almost no wall-clock and does
        // buy a dense sheet time to finish drawing.
        "--virtual-time-budget=60000",
        ...extra,
        `http://127.0.0.1:${PORT}/index.html`,
      ], { maxBuffer: 64 * 1024 * 1024, timeout: 120_000, killSignal: "SIGKILL" });

      return stdout;
    };

    // First: how many pages, and how big. Nothing is drawn, so this run is
    // small whatever the document is.
    const survey = await draw(0, ["--window-size=800,600", "--dump-dom"]);

    const failure = /<pre id="raster-error">([\s\S]*?)<\/pre>/.exec(survey);
    if (failure !== null) {
      throw new Error(`pdf.js: ${failure[1].slice(0, 300)}`);
    }

    const shape = /<div id="raster-done">(\d+)x(\d+)x(\d+)<\/div>/.exec(survey);
    if (shape === null) {
      throw new Error("the raster did not finish");
    }

    const total = Number(shape[1]);
    const width = Number(shape[2]);
    const height = Number(shape[3]);
    const wanted = Math.min(total, PAGE_CAP);

    // Said out loud, because a cap nobody is told about reads as "that was all
    // there was" — and on this suite the document that overruns its pages is
    // the one whose page count is the finding.
    if (wanted < total) {
      console.log(`  raster: ${wanted} of ${total} pages (capped at ${PAGE_CAP})`);
    }

    const written = [];

    for (let page = 1; page <= wanted; page += 1) {
      const file = `word-p${page}.png`;
      await draw(page, [`--window-size=${width},${height}`, `--screenshot=${resolve(outDir, file)}`]);
      written.push({ file });
    }

    return written;
  } finally {
    if (server) {
      server.closeAllConnections?.();
      await new Promise((done) => server.close(done));
    }
    if (tempDir) {
      await rm(tempDir, { recursive: true, force: true }).catch(() => {});
    }
  }
}
