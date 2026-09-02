/**
 * Probe V — pixels.
 *
 * Rasterises the Word-exported `.verify/invoice.pdf` (probe C's side-product)
 * at 96dpi and pixel-compares every visual region against the design fixtures
 * `design-p1.png` / `design-p2.png`. Writes `.verify/measure-v.json`.
 *
 * The rasterising happens in headless Chrome, not Node — Node has no canvas,
 * and shipping a native canvas build for one probe is not worth it. A tiny
 * same-origin server (port 8901, per the contract) serves a harness page, the
 * PDF and pdfjs-dist out of node_modules; the harness renders each page to a
 * canvas at scale 96/72 (A4 -> ~794x1123), writes `toDataURL()` into the DOM,
 * and `--dump-dom` carries the pixels back to Node where pngjs decodes them.
 *
 * Region rects come from `design/invoice-v2/fixtures/design-regions.json`
 * (px @ 96dpi, page-relative). Two masks/exemptions, per the contract:
 *  - `summary-label`: the design draws a non-printing annotation chip right of
 *    the label (plan D12), so everything beyond 60% of the region's width is
 *    painted page-white on BOTH crops before diffing.
 *  - `scan-card` / `qr-canvas`: geometry-only (the canvas QR is random
 *    modules); recorded as skipped, never pixel-compared.
 *
 * If probe B left `preview-p1.png` / `preview-p2.png` behind, the same regions
 * are also diffed preview-vs-Word (`crossEngine`). No chip mask there: the
 * chip is a design-canvas annotation and neither engine prints it.
 *
 * This probe never runs Word itself — a missing PDF is recorded as
 * `{ ok: false, reason: "no pdf" }`, not worked around.
 */
import { createServer } from "node:http";
import { spawn } from "node:child_process";
import { readFile, writeFile } from "node:fs/promises";
import { existsSync, readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import pixelmatch from "pixelmatch";
import { PNG } from "pngjs";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..", "..");
const VERIFY_DIR = resolve(ROOT, ".verify");
const FIXTURE_DIR = resolve(ROOT, "design", "invoice-v2", "fixtures");
const PDFJS_DIR = resolve(ROOT, "node_modules", "pdfjs-dist", "build");
const CHROME = "C:/Program Files/Google/Chrome/Application/chrome.exe";
const PORT = 8901;

/** Regions with `visual: yes` in the contract's region table. */
export const VISUAL_REGIONS = [
  { region: "letterhead", page: 1 },
  { region: "rule", page: 1 },
  { region: "band", page: 1 },
  { region: "status-pill", page: 1 },
  { region: "parties", page: 1 },
  { region: "summary-label", page: 1 },
  { region: "summary", page: 1 },
  { region: "charges-head", page: 1 },
  { region: "charges-body", page: 1 },
  { region: "totals-panel", page: 1 },
  { region: "total-bar", page: 1 },
  { region: "closer", page: 1 },
  { region: "footer", page: 1 },
  { region: "p2-letterhead", page: 2 },
  { region: "bank-grid", page: 2 },
  { region: "reference-panel", page: 2 },
  { region: "terms", page: 2 },
  { region: "footer-2", page: 2 },
];

/** Geometry-only regions: recorded, never pixel-compared. */
export const GEOMETRY_ONLY_REGIONS = [
  { region: "scan-card", page: 2 },
  { region: "qr-canvas", page: 2 },
];

/** Fraction of `summary-label`'s width kept; the rest hides the design chip. */
const SUMMARY_LABEL_KEEP = 0.6;

export function readPng(path) {
  return PNG.sync.read(readFileSync(path));
}

/**
 * Integer rect clamped to an image, or null when nothing of it is inside.
 * Rects in the fixture are fractional px; the crop has to be whole pixels.
 */
function clampRect(rect, width, height) {
  const x = Math.max(0, Math.round(rect.x));
  const y = Math.max(0, Math.round(rect.y));
  const w = Math.min(Math.round(rect.w), width - x);
  const h = Math.min(Math.round(rect.h), height - y);
  return w > 0 && h > 0 ? { x, y, w, h } : null;
}

/** Copies a rect out of a PNG into a fresh one, at exactly w x h. */
function cropPng(png, rect) {
  const out = new PNG({ width: rect.w, height: rect.h });
  for (let row = 0; row < rect.h; row += 1) {
    const src = ((rect.y + row) * png.width + rect.x) * 4;
    png.data.copy(out.data, row * rect.w * 4, src, src + rect.w * 4);
  }
  return out;
}

/** Paints every column from `fraction` of the width rightwards page-white. */
function maskBeyondFraction(png, fraction) {
  const from = Math.floor(png.width * fraction);
  for (let y = 0; y < png.height; y += 1) {
    for (let x = from; x < png.width; x += 1) {
      const i = (y * png.width + x) * 4;
      png.data[i] = 255;
      png.data[i + 1] = 255;
      png.data[i + 2] = 255;
      png.data[i + 3] = 255;
    }
  }
}

/**
 * Diffs one region between a reference image (whose space the rect lives in)
 * and another render of the same page. The rect is scaled proportionally into
 * the other image's space — the raster is expected at 794x1123 but a pixel or
 * two of rounding must not shear every crop — then both crops are taken at
 * the common (smallest) dimensions so pixelmatch sees equal buffers.
 *
 * @returns {{ pctDiff: number, diffPixels: number, totalPixels: number } | null}
 *   null when the rect falls outside either image.
 */
export function diffRegion(referencePng, otherPng, rect, { maskKeepFraction } = {}) {
  const refRect = clampRect(rect, referencePng.width, referencePng.height);
  const scaled = {
    x: rect.x * (otherPng.width / referencePng.width),
    y: rect.y * (otherPng.height / referencePng.height),
    w: rect.w * (otherPng.width / referencePng.width),
    h: rect.h * (otherPng.height / referencePng.height),
  };
  const otherRect = clampRect(scaled, otherPng.width, otherPng.height);
  if (!refRect || !otherRect) return null;

  const w = Math.min(refRect.w, otherRect.w);
  const h = Math.min(refRect.h, otherRect.h);
  const a = cropPng(referencePng, { ...refRect, w, h });
  const b = cropPng(otherPng, { ...otherRect, w, h });

  if (typeof maskKeepFraction === "number") {
    // Same mask on both sides, so the masked area contributes zero diffs.
    maskBeyondFraction(a, maskKeepFraction);
    maskBeyondFraction(b, maskKeepFraction);
  }

  const diffPixels = pixelmatch(a.data, b.data, null, w, h, { threshold: 0.1 });
  const totalPixels = w * h;
  return {
    pctDiff: Number(((diffPixels / totalPixels) * 100).toFixed(2)),
    diffPixels,
    totalPixels,
  };
}

/** The six regions the design draws with rounded corners — contract F15. */
export const ROUNDED_REGIONS = [
  { region: "status-pill", page: 1 },
  { region: "charges-head", page: 1 },
  { region: "totals-panel", page: 1 },
  { region: "total-bar", page: 1 },
  { region: "reference-panel", page: 2 },
  { region: "scan-card", page: 2 },
];

/** How far apart two colours must be before they count as different ink. */
const INK_DISTANCE = 24;

/** The five entities a DOM dump escapes, undone. */
function decodeHtml(text) {
  return text
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&");
}

/** One pixel as [r,g,b], or null outside the image. */
function pixelAt(png, x, y) {
  const px = Math.round(x);
  const py = Math.round(y);
  if (px < 0 || py < 0 || px >= png.width || py >= png.height) return null;
  const i = (py * png.width + px) * 4;
  return [png.data[i], png.data[i + 1], png.data[i + 2]];
}

/** Manhattan distance between two colours — cheap, and enough to tell ink apart. */
function inkDistance(a, b) {
  if (!a || !b) return Infinity;
  return Math.abs(a[0] - b[0]) + Math.abs(a[1] - b[1]) + Math.abs(a[2] - b[2]);
}

/**
 * The corner test for one region, in one render.
 *
 * A rounded box does not paint its own bounding-box corners: 2px in along the
 * diagonal still shows whatever was behind it. A square box paints them the
 * same as its middle. So the test is corner-against-centre, with the ground
 * sampled just outside the box to say which of the two a corner looks like.
 *
 * A region with no fill of its own (the scan card is an outline on white) has
 * no centre to differ from, and this test cannot speak to it — that is
 * recorded as `applicable: false` rather than guessed either way.
 */
function cornerTest(png, rect, region) {
  const fill = pixelAt(png, rect.x + rect.w / 2, rect.y + rect.h / 2);
  // Ground: outside each corner on the same diagonal, averaged by majority —
  // one sample could land on a neighbouring element.
  const outside = [
    pixelAt(png, rect.x - 3, rect.y - 3),
    pixelAt(png, rect.x + rect.w + 2, rect.y - 3),
    pixelAt(png, rect.x - 3, rect.y + rect.h + 2),
    pixelAt(png, rect.x + rect.w + 2, rect.y + rect.h + 2),
  ].filter(Boolean);
  const ground = outside[0] ?? null;

  const applicable = fill !== null && ground !== null && inkDistance(fill, ground) > INK_DISTANCE;
  const inset = 2;
  const corners = [
    pixelAt(png, rect.x + inset, rect.y + inset),
    pixelAt(png, rect.x + rect.w - 1 - inset, rect.y + inset),
    pixelAt(png, rect.x + inset, rect.y + rect.h - 1 - inset),
    pixelAt(png, rect.x + rect.w - 1 - inset, rect.y + rect.h - 1 - inset),
  ].map((corner) => inkDistance(corner, fill) > inkDistance(corner, ground));

  return {
    region,
    applicable,
    // A region that cannot be judged reports no verdict rather than a false one.
    corners: applicable ? corners : null,
    centerOk: applicable,
    fill,
    ground,
  };
}

/**
 * The corner test across every rounded region of one render, with the design
 * rects scaled into that render's space the same way `diffRegion` scales them.
 */
export function cornerTestsFor(pages, rectsByPage, referencePages) {
  const out = [];

  for (const { region, page } of ROUNDED_REGIONS) {
    const rect = rectsByPage[page - 1]?.[region];
    const png = pages[page - 1];
    const reference = referencePages[page - 1];
    if (!rect || !png || !reference) continue;

    const sx = png.width / reference.width;
    const sy = png.height / reference.height;
    out.push(cornerTest(png, {
      x: rect.x * sx, y: rect.y * sy, w: rect.w * sx, h: rect.h * sy,
    }, region));
  }

  return out;
}

/**
 * The page pdfjs runs in. Everything it needs is served same-origin.
 *
 * Two timing tricks, both earned the hard way:
 *  - `--dump-dom` dumps at the window load event and will not wait for a
 *    module script's top-level await. The `<img src="/hold.png">` holds the
 *    load event open — the server withholds that response until the harness
 *    calls `/release` — so the dump always contains the finished pages.
 *    (`--virtual-time-budget` is no substitute: its virtual clock hits
 *    Chrome's task-starvation guard mid-render and dumps 2 of 3 pages,
 *    however large the budget.)
 *  - The worker module is imported on the main thread (pdfjs's "fake worker"
 *    path, keyed off `globalThis.pdfjsWorker`), so the whole render is one
 *    thread and one event loop — nothing for the dump to race against.
 *
 * The data URLs land in `<script type="pv-data">` holders: script bodies are
 * serialised by dump-dom but never laid out, so a couple of MB of base64
 * costs the renderer nothing.
 */
function harnessHtml(pdfUrl) {
  return `<!doctype html>
<html><head><meta charset="utf-8"><title>probe-v raster</title></head><body>
<img src="/hold.png" alt="">
<script type="module">
  try {
    const pdfjs = await import("/pdfjs/pdf.min.mjs");
    globalThis.pdfjsWorker = await import("/pdfjs/pdf.worker.min.mjs");
    pdfjs.GlobalWorkerOptions.workerSrc = "/pdfjs/pdf.worker.min.mjs";
    const doc = await pdfjs.getDocument({ url: new URL("${pdfUrl}", location.href).href }).promise;
    for (let n = 1; n <= doc.numPages; n += 1) {
      const page = await doc.getPage(n);
      const viewport = page.getViewport({ scale: 96 / 72 });
      const canvas = document.createElement("canvas");
      canvas.width = Math.round(viewport.width);
      canvas.height = Math.round(viewport.height);
      await page.render({ canvas, viewport }).promise;
      const holder = document.createElement("script");
      holder.type = "pv-data";
      holder.id = "page-" + n;
      holder.textContent = canvas.toDataURL("image/png");
      document.body.appendChild(holder);

      // Word evaluates PAGE once per story, so "1 / 2" then "2 / 2" exists
      // nowhere but the rendered pages — which makes the PDF the only place
      // C6 can be checked. The strip is found by its lowest text, whatever
      // height the footer happens to be: a band fixed as a fraction of the
      // page catches body copy on a full page and nothing on a short one.
      const text = await page.getTextContent();
      const placed = text.items.filter((item) => item.transform && item.str.trim() !== "");
      const lowest = placed.reduce((min, item) => Math.min(min, item.transform[5]), Infinity);
      const strip = placed
        .filter((item) => item.transform[5] < lowest + 12)
        .sort((a, b) => a.transform[4] - b.transform[4])
        .map((item) => item.str)
        .join(" ")
        .replace(/\\s+/g, " ")
        .trim();
      const textHolder = document.createElement("script");
      textHolder.type = "pv-data";
      textHolder.id = "footer-text-" + n;
      textHolder.textContent = strip;
      document.body.appendChild(textHolder);
    }
    document.body.appendChild(Object.assign(document.createElement("div"), { id: "raster-done" }));
  } catch (error) {
    const pre = document.createElement("pre");
    pre.id = "raster-error";
    pre.textContent = String(error && error.stack || error);
    document.body.appendChild(pre);
  }
  fetch("/release").catch(() => {});
</script>
</body></html>`;
}

/** 1x1 transparent PNG, the body of the load-holding image. */
const HOLD_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==",
  "base64",
);

/**
 * Starts the static server. Retries on EADDRINUSE — the contract pins port
 * 8901 for every probe, and a concurrently running probe B may hold it for a
 * moment; waiting beats racing to a different port the contract forbids.
 */
function startServer(routes, { retries = 30, retryMs = 2000 } = {}) {
  const heldResponses = [];
  const server = createServer((req, res) => {
    const path = new URL(req.url, `http://127.0.0.1:${PORT}`).pathname;
    if (path === "/hold.png") {
      // Withheld until /release: this open request keeps the page's load
      // event — and with it Chrome's --dump-dom — waiting for the harness.
      heldResponses.push(res);
      return;
    }
    if (path === "/release") {
      for (const held of heldResponses.splice(0)) {
        held.writeHead(200, { "content-type": "image/png" }).end(HOLD_PNG);
      }
      res.writeHead(200, { "content-type": "text/plain" }).end("ok");
      return;
    }
    const route = routes[path];
    if (!route) {
      res.writeHead(404).end("not found");
      return;
    }
    try {
      const body = route.body ?? readFileSync(route.file);
      res.writeHead(200, { "content-type": route.type }).end(body);
    } catch (error) {
      res.writeHead(500).end(String(error));
    }
  });

  return new Promise((resolveStart, rejectStart) => {
    let attempts = 0;
    const listen = () => server.listen(PORT, "127.0.0.1", () => resolveStart(server));
    server.on("error", (error) => {
      if (error.code === "EADDRINUSE" && attempts < retries) {
        attempts += 1;
        setTimeout(listen, retryMs);
      } else {
        rejectStart(error);
      }
    });
    listen();
  });
}

/** Runs Chrome headless with --dump-dom and returns the serialised DOM. */
function dumpDom(url, { timeoutMs = 120000 } = {}) {
  return new Promise((resolveDump, rejectDump) => {
    const chrome = spawn(CHROME, [
      "--headless=new",
      "--disable-gpu",
      "--hide-scrollbars",
      "--window-size=900,1200",
      // No --virtual-time-budget here, deliberately: the harness holds the
      // load event open until rendering is done (see harnessHtml), which is
      // what actually makes --dump-dom wait. Virtual time was measured to
      // trip Chrome's task-starvation guard mid-render and dump early at any
      // budget; the Node-side timer below is the hang guard instead.
      "--dump-dom",
      url,
    ], { stdio: ["ignore", "pipe", "pipe"] });

    const out = [];
    const err = [];
    chrome.stdout.on("data", (chunk) => out.push(chunk));
    chrome.stderr.on("data", (chunk) => err.push(chunk));

    // dump-dom exits by itself; the timer only guards a hung render.
    const timer = setTimeout(() => {
      chrome.kill("SIGKILL");
      rejectDump(new Error(`chrome timed out after ${timeoutMs}ms`));
    }, timeoutMs);

    chrome.on("error", (error) => {
      clearTimeout(timer);
      rejectDump(error);
    });
    chrome.on("close", (code) => {
      clearTimeout(timer);
      const dom = Buffer.concat(out).toString("utf8");
      if (dom.length > 0) resolveDump(dom);
      else rejectDump(new Error(`chrome exited ${code} with no DOM: ${Buffer.concat(err).toString("utf8").slice(0, 500)}`));
    });
  });
}

/**
 * Rasterises a PDF to one PNG per page, via the served harness.
 * @returns {Promise<PNG[]>} decoded pages, in order.
 */
export async function rasterisePdf(pdfPath) {
  const routes = {
    "/harness.html": { body: harnessHtml("/invoice.pdf"), type: "text/html; charset=utf-8" },
    "/invoice.pdf": { file: pdfPath, type: "application/pdf" },
    "/pdfjs/pdf.min.mjs": { file: resolve(PDFJS_DIR, "pdf.min.mjs"), type: "text/javascript" },
    "/pdfjs/pdf.worker.min.mjs": { file: resolve(PDFJS_DIR, "pdf.worker.min.mjs"), type: "text/javascript" },
  };

  const server = await startServer(routes);
  try {
    const dom = await dumpDom(`http://127.0.0.1:${PORT}/harness.html`);

    const errorMatch = dom.match(/<pre id="raster-error">([\s\S]*?)<\/pre>/);
    if (errorMatch) throw new Error(`pdfjs harness failed: ${errorMatch[1].slice(0, 500)}`);
    if (!dom.includes('id="raster-done"')) throw new Error("raster incomplete: no done marker in DOM");

    const pages = [];
    const pattern = /<script type="pv-data" id="page-(\d+)">data:image\/png;base64,([A-Za-z0-9+/=]+)<\/script>/g;
    for (const match of dom.matchAll(pattern)) {
      pages[Number(match[1]) - 1] = PNG.sync.read(Buffer.from(match[2], "base64"));
    }
    if (pages.length === 0) throw new Error("raster produced no pages");

    // Per-page footer text, which only exists once the fields are rendered.
    const footerTextByPage = [];
    const textPattern = /<script type="pv-data" id="footer-text-(\d+)">([\s\S]*?)<\/script>/g;
    for (const match of dom.matchAll(textPattern)) {
      footerTextByPage[Number(match[1]) - 1] = decodeHtml(match[2]).trim();
    }

    return { pages, footerTextByPage };
  } finally {
    // Success or failure, the port goes back — the runner reuses it. Chrome's
    // keep-alive sockets would otherwise hold close() open for their timeout.
    server.closeAllConnections?.();
    await new Promise((done) => server.close(done));
  }
}

/**
 * Diffs every visual region of a two-page render against the design fixtures.
 * Exported so the self-test can run the identical path fixture-vs-fixture.
 *
 * @param {PNG[]} designPages [design-p1, design-p2]
 * @param {PNG[]} renderPages the pages under test, same order
 * @param {Record<string, any>[]} rectsByPage [page1 rects, page2 rects]
 * @param {{ chipMask?: boolean }} options chip mask on summary-label (design side only carries the chip)
 */
export function diffAllRegions(designPages, renderPages, rectsByPage, { chipMask = true } = {}) {
  const results = [];

  for (const { region, page } of VISUAL_REGIONS) {
    const rect = rectsByPage[page - 1]?.[region];
    const designPng = designPages[page - 1];
    const renderPng = renderPages[page - 1];
    if (!rect || !designPng || !renderPng) {
      results.push({ region, page, pctDiff: null, diffPixels: null, totalPixels: null, note: rect ? "page missing" : "no rect" });
      continue;
    }
    const mask = chipMask && region === "summary-label" ? { maskKeepFraction: SUMMARY_LABEL_KEEP } : {};
    const diff = diffRegion(designPng, renderPng, rect, mask);
    if (!diff) {
      results.push({ region, page, pctDiff: null, diffPixels: null, totalPixels: null, note: "rect outside image" });
      continue;
    }
    results.push({ region, page, ...diff });
  }

  for (const { region, page } of GEOMETRY_ONLY_REGIONS) {
    results.push({ region, page, skipped: "geometry-only" });
  }

  return results;
}

/** Loads the fixture rects as [page1, page2] maps. */
async function loadRects() {
  const regions = JSON.parse(await readFile(resolve(FIXTURE_DIR, "design-regions.json"), "utf8"));
  return [regions.page1, regions.page2];
}

/**
 * The probe. Reads `.verify/invoice.pdf`, writes `.verify/measure-v.json`,
 * and returns what it wrote. Never throws: every failure mode lands in the
 * JSON as `{ ok: false, reason }` so the runner can put it on the board.
 */
export async function runVisualProbe() {
  const outPath = resolve(VERIFY_DIR, "measure-v.json");
  const write = async (measure) => {
    await writeFile(outPath, JSON.stringify(measure, null, 2), "utf8");
    return measure;
  };

  const pdfPath = resolve(VERIFY_DIR, "invoice.pdf");
  if (!existsSync(pdfPath)) {
    return write({ ok: false, reason: "no pdf", rasterPages: [], regions: [], crossEngine: [] });
  }

  let designPages;
  let rectsByPage;
  try {
    designPages = [readPng(resolve(FIXTURE_DIR, "design-p1.png")), readPng(resolve(FIXTURE_DIR, "design-p2.png"))];
    rectsByPage = await loadRects();
  } catch (error) {
    return write({ ok: false, reason: `fixtures unreadable: ${error.message}`, rasterPages: [], regions: [], crossEngine: [] });
  }

  let rasterPages;
  let footerTextByPage = [];
  try {
    const rastered = await rasterisePdf(pdfPath);
    rasterPages = rastered.pages;
    footerTextByPage = rastered.footerTextByPage;
  } catch (error) {
    return write({ ok: false, reason: `raster failed: ${error.message}`, rasterPages: [], regions: [], crossEngine: [] });
  }

  const regions = diffAllRegions(designPages, rasterPages, rectsByPage);

  // Cross-engine (probe B's screenshots vs the Word raster): informational
  // until gate G4, so their absence is not a failure — just an empty list.
  const crossEngine = [];
  const previewPaths = [resolve(VERIFY_DIR, "preview-p1.png"), resolve(VERIFY_DIR, "preview-p2.png")];
  let previewPages = null;
  if (previewPaths.every((path) => existsSync(path))) {
    previewPages = previewPaths.map((path) => readPng(path));
    // The rects live in design-fixture space; diffRegion scales them into the
    // preview's space, then the preview crop is compared against the Word
    // raster crop. No chip mask: neither engine prints the design chip.
    for (const entry of diffAllRegions(previewPages, rasterPages, rectsByPage, { chipMask: false })) {
      crossEngine.push(entry);
    }
  }

  // F15 asks the same question of both renders: does anything actually draw a
  // rounded corner? The design fixture is the rect source for both.
  const cornerTests = {
    word: cornerTestsFor(rasterPages, rectsByPage, designPages),
    preview: previewPages === null ? null : cornerTestsFor(previewPages, rectsByPage, designPages),
  };

  return write({
    ok: true,
    rasterPages: rasterPages.map((page) => ({ w: page.width, h: page.height })),
    regions,
    crossEngine,
    cornerTests,
    footerTextByPage,
  });
}

// CLI: node scripts/lib/probe-visual.mjs
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  runVisualProbe().then((measure) => {
    if (!measure.ok) {
      console.log(`probe V: ok=false (${measure.reason})`);
      return;
    }
    console.log("raster pages:", measure.rasterPages.map(({ w, h }) => `${w}x${h}`).join(", "));
    for (const entry of measure.regions) {
      if (entry.skipped) console.log(`  ${entry.region.padEnd(16)} skipped (${entry.skipped})`);
      else if (entry.pctDiff === null) console.log(`  ${entry.region.padEnd(16)} null (${entry.note})`);
      else console.log(`  ${entry.region.padEnd(16)} ${String(entry.pctDiff).padStart(6)}%  (${entry.diffPixels}/${entry.totalPixels})`);
    }
    if (measure.crossEngine.length > 0) console.log(`cross-engine entries: ${measure.crossEngine.length}`);
  }, (error) => {
    console.error("probe V crashed:", error);
    process.exitCode = 1;
  });
}
