/**
 * Puts a design canvas and the preview of the document built from it side by
 * side, page by page, and scores how alike they are.
 *
 *   node scripts/design-parity.mjs portfolio-report
 *   node scripts/design-parity.mjs portfolio-report --pages 1,3
 *
 * Reads `design/<id>/<Design>.dc.html` (the canvas, rendered by its own
 * support.js) and `public/demo/<id>.html` (the packed .docx, laid out by
 * docx-preview — run `npm run demo` first). Both are photographed in the same
 * headless Chrome at 96dpi, so one A4 page is 794×1123 on each side.
 *
 * Writes `.parity/<id>/`: design-NN.png, preview-NN.png, diff-NN.png,
 * scores.json, and index.html — the board, which shows each page's pair, the
 * difference, an overlay with a slider, and the score.
 *
 * Two scores per page, because each alone misleads:
 *  - pixel: share of pixels pixelmatch calls the same. Harsh — a line that
 *    wraps one word early repaints a whole paragraph — but it is what "100%"
 *    has to mean in the end.
 *  - layout: the two pages shrunk to 1/8 size, grey, and compared by mean
 *    difference. Forgives glyph rasterising and one-pixel drift; punishes a
 *    block in the wrong place, the wrong fill, or a missing chart.
 * The headline is their mean.
 *
 * @module
 */

import { createServer } from "node:http";
import { mkdir, readFile, readdir, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { dirname, extname, join, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";
import { PNG } from "pngjs";
import pixelmatch from "pixelmatch";
import { capture, openChrome } from "./lib/cdp.mjs";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const TYPES = {
  ".html": "text/html",
  ".js": "text/javascript",
  ".mjs": "text/javascript",
  ".css": "text/css",
  ".json": "application/json",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".webp": "image/webp",
  ".jpg": "image/jpeg",
  ".pdf": "application/pdf",
};

/**
 * Rewrites the canvas's prop defaults, which is the only input it takes.
 *
 * The canvas has editor props — `mode` (engine or preview), `highlightAi` —
 * that the design tool sets from its toolbar. Headless there is no toolbar,
 * so the defaults in `data-props` are what it renders with.
 */
function withProps(html, props) {
  return html.replace(/data-props="([^"]*)"/, (_, encoded) => {
    const spec = JSON.parse(encoded.replaceAll("&quot;", '"').replaceAll("&amp;", "&"));
    for (const [key, value] of Object.entries(props)) if (spec[key]) spec[key].default = value;
    return `data-props="${JSON.stringify(spec).replaceAll("&", "&amp;").replaceAll('"', "&quot;")}"`;
  });
}

/** Serves website/ so the canvas and the preview load from one origin. */
function serve(canvasProps) {
  const server = createServer(async (req, res) => {
    const url = new URL(req.url, "http://x");
    if (url.pathname === "/__pdf") {
      res.writeHead(200, { "content-type": "text/html" });
      return res.end(pdfPage(url.searchParams.get("file")));
    }
    const path = resolve(ROOT, "." + decodeURIComponent(url.pathname));
    if (!path.startsWith(ROOT + sep) || !existsSync(path)) {
      res.writeHead(404);
      return res.end();
    }
    let body = await readFile(path);
    if (path.endsWith(".dc.html")) body = withProps(body.toString("utf8"), canvasProps);
    res.writeHead(200, { "content-type": TYPES[extname(path)] ?? "application/octet-stream" });
    res.end(body);
  });
  return new Promise((ok) => server.listen(0, "127.0.0.1", () => ok(server)));
}

/** Waits until `selector` matches `count` elements and every image and font has settled. */
const settle = (selector, minimum) => `(async () => {
  for (let i = 0; i < 300; i++) {
    if (document.querySelectorAll(${JSON.stringify(selector)}).length >= ${minimum}) break;
    await new Promise((r) => setTimeout(r, 100));
  }
  for (let i = 0; i < 600 && location.pathname === "/__pdf" && !document.body.dataset.done; i++) {
    await new Promise((r) => setTimeout(r, 100));
  }
  await document.fonts.ready;
  await Promise.all([...document.images].map((img) => img.complete ? null : new Promise((r) => { img.onload = img.onerror = r; })));
  await new Promise((r) => setTimeout(r, 500));
  // Inside the border: the canvas rules a hairline round a sheet to show its
  // edge on screen, which is chrome, not page — measured from outside it,
  // every word on the page reads a pixel out.
  return [...document.querySelectorAll(${JSON.stringify(selector)})].map((el) => {
    const r = el.getBoundingClientRect();
    const cs = getComputedStyle(el);
    const [t, rr, b, l] = ["Top", "Right", "Bottom", "Left"].map((side) => parseFloat(cs["border" + side + "Width"]) || 0);
    return { x: r.x + scrollX + l, y: r.y + scrollY + t, width: r.width - l - rr, height: r.height - t - b };
  });
})()`;

/**
 * The face a stack actually draws in on this machine: the first one installed.
 * The canvas names Aptos and Cambria first, which this Mac does not have.
 */
const face = `(() => {
  const ctx = document.createElement("canvas").getContext("2d");
  const width = (font) => { ctx.font = "40px " + font; return ctx.measureText("mmmmmlllllWWWii0123").width; };
  const installed = (name) => ["monospace", "serif"].some((fallback) => width(JSON.stringify(name) + "," + fallback) !== width(fallback));
  const generic = new Set(["serif", "sans-serif", "monospace", "cursive", "system-ui", "ui-monospace"]);
  return (stack) => {
    for (const name of stack.split(",").map((f) => f.trim().replace(/["']/g, ""))) {
      if (generic.has(name) || installed(name)) return name;
    }
    return stack;
  };
})()`;

/**
 * Every run of text on each page: what it says, where it sits from the
 * page's corner, and how it is set. Read from text nodes rather than elements,
 * so a span in the canvas and a run in the preview meet on the words.
 */
const readTexts = (selector) => `(() => { const face = ${face}; return [...document.querySelectorAll(${JSON.stringify(selector)})].map((sheet) => {
  const box = sheet.getBoundingClientRect();
  const sheetStyle = getComputedStyle(sheet);
  const origin = { x: box.x + (parseFloat(sheetStyle.borderLeftWidth) || 0), y: box.y + (parseFloat(sheetStyle.borderTopWidth) || 0) };
  const walker = document.createTreeWalker(sheet, NodeFilter.SHOW_TEXT);
  const out = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const raw = node.textContent.replace(/\\s+/g, " ").trim();
    if (!raw) continue;
    const el = node.parentElement;
    const cs = getComputedStyle(el);
    // What the reader sees: a run set in capitals by its style reads as capitals.
    const text = cs.textTransform === "uppercase" ? raw.toUpperCase() : raw;
    if (cs.visibility === "hidden" || cs.display === "none") continue;
    const range = document.createRange();
    range.selectNodeContents(node);
    const r = range.getClientRects()[0];
    if (!r || r.width === 0) continue;
    out.push({
      text,
      x: +(r.x - origin.x).toFixed(1),
      y: +(r.y - origin.y).toFixed(1),
      w: +r.width.toFixed(1),
      h: +r.height.toFixed(1),
      size: cs.fontSize,
      family: face(cs.fontFamily),
      weight: cs.fontWeight,
      color: cs.color,
      spacing: cs.letterSpacing,
    });
  }
  return out;
}); })()`;

/**
 * Both sides rasterised alike. The canvas's design system sets its body
 * `-webkit-font-smoothing: antialiased`; a page without it draws the same
 * glyphs heavier, and every word would differ by its edge pixels alone. That
 * is how this screen draws type, not something either document says.
 */
const SAME_SMOOTHING = `(() => { const style = document.createElement("style");
  style.textContent = "* { -webkit-font-smoothing: antialiased !important; }";
  document.head.append(style); })()`;

/** Pages as rects, from the canvas or from the preview, and the text on each. */
async function photograph(page, url, selector, minimum) {
  await page.goto(url);
  await page.evaluate(SAME_SMOOTHING);
  const rects = await page.evaluate(settle(selector, minimum));
  const failed = await page.evaluate("document.body.dataset.error ?? null");
  if (failed) throw new Error(`${url} failed: ${failed}`);
  const texts = await page.evaluate(readTexts(selector));
  const shots = [];
  for (const rect of rects) shots.push(PNG.sync.read(await capture(page, rect)));
  return { shots, texts };
}

/**
 * The canvas's words against the preview's, page by page: each string found on
 * both, matched in reading order, and how far it moved or how it is set
 * differently. Strings on one side only are listed too — a word the preview
 * dropped, or a run it split differently.
 */
function textDrift(designTexts, previewTexts) {
  const pool = new Map();
  for (const t of previewTexts) {
    if (!pool.has(t.text)) pool.set(t.text, []);
    pool.get(t.text).push(t);
  }
  const moved = [];
  const missing = [];
  for (const d of designTexts) {
    const candidates = pool.get(d.text);
    const p = candidates?.shift();
    if (!p) {
      missing.push(d.text);
      continue;
    }
    const style = ["size", "family", "weight", "color", "spacing"].filter((k) => d[k] !== p[k]).map((k) => `${k} ${d[k]} → ${p[k]}`);
    const dx = +(p.x - d.x).toFixed(1);
    const dy = +(p.y - d.y).toFixed(1);
    if (Math.abs(dx) > 1 || Math.abs(dy) > 1 || style.length > 0) moved.push({ text: d.text.slice(0, 48), dx, dy, style });
  }
  const extra = [...pool.values()].flat().map((t) => t.text);
  return { moved, missing, extra };
}

/**
 * The pages of a PDF — Word's own export of the packed file — drawn by pdf.js
 * at 96dpi on canvases, so they can be photographed like any other page.
 */
const pdfPage = (pdfUrl) => `<!doctype html><html><body style="margin:0;background:#fff">
<script>addEventListener("error", (e) => { document.body.dataset.done = "1"; document.body.dataset.error = e.message; });
addEventListener("unhandledrejection", (e) => { document.body.dataset.done = "1"; document.body.dataset.error = String(e.reason); });</script>
<script type="module">
import * as pdfjs from "/node_modules/pdfjs-dist/build/pdf.mjs";
pdfjs.GlobalWorkerOptions.workerSrc = "/node_modules/pdfjs-dist/build/pdf.worker.mjs";
const pdf = await pdfjs.getDocument({ url: ${JSON.stringify(pdfUrl)} }).promise;
for (let n = 1; n <= pdf.numPages; n++) {
  const page = await pdf.getPage(n);
  const viewport = page.getViewport({ scale: 96 / 72 });
  const canvas = document.createElement("canvas");
  canvas.className = "pdf-page";
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  canvas.style.display = "block";
  document.body.append(canvas);
  await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
}
document.body.dataset.done = "1";
</script></body></html>`;

/** Crops or pads `png` onto a white canvas of `width`×`height`. */
function fit(png, width, height) {
  const out = new PNG({ width, height });
  out.data.fill(255);
  PNG.bitblt(png, out, 0, 0, Math.min(width, png.width), Math.min(height, png.height), 0, 0);
  return out;
}

/** The design and the render side by side at half size, for reading at a glance. */
function pair(design, render) {
  const half = (png) => {
    const out = new PNG({ width: png.width >> 1, height: png.height >> 1 });
    for (let y = 0; y < out.height; y++) {
      for (let x = 0; x < out.width; x++) {
        for (let c = 0; c < 4; c++) out.data[(y * out.width + x) * 4 + c] = png.data[(2 * y * png.width + 2 * x) * 4 + c];
      }
    }
    return out;
  };
  const a = half(design);
  const b = half(fit(render, design.width, design.height));
  const out = new PNG({ width: a.width * 2 + 8, height: a.height });
  out.data.fill(255);
  PNG.bitblt(a, out, 0, 0, a.width, a.height, 0, 0);
  PNG.bitblt(b, out, 0, 0, b.width, b.height, a.width + 8, 0);
  return out;
}

/** Mean grey difference of two pages shrunk by `factor`, as a similarity in [0, 1]. */
function layoutScore(a, b, factor = 8) {
  const w = Math.floor(a.width / factor);
  const h = Math.floor(a.height / factor);
  const grey = (png) => {
    const cells = new Float64Array(w * h);
    for (let y = 0; y < h * factor; y++) {
      for (let x = 0; x < w * factor; x++) {
        const i = (y * png.width + x) * 4;
        cells[Math.floor(y / factor) * w + Math.floor(x / factor)] +=
          0.299 * png.data[i] + 0.587 * png.data[i + 1] + 0.114 * png.data[i + 2];
      }
    }
    return cells.map((sum) => sum / (factor * factor));
  };
  const ga = grey(a);
  const gb = grey(b);
  let total = 0;
  for (let i = 0; i < ga.length; i++) total += Math.abs(ga[i] - gb[i]);
  return 1 - total / ga.length / 255;
}

function compare(design, preview) {
  const width = design.width;
  const height = design.height;
  const b = fit(preview, width, height);
  const diff = new PNG({ width, height });
  const changed = pixelmatch(design.data, b.data, diff.data, width, height, { threshold: 0.15, alpha: 0.2 });
  const pixel = 1 - changed / (width * height);
  const layout = layoutScore(design, b);
  return { diff, pixel, layout, score: (pixel + layout) / 2 };
}

const percent = (n) => `${(n * 100).toFixed(1)}%`;

function board(id, results) {
  const overall = results.reduce((sum, r) => sum + r.score, 0) / results.length;
  const rows = results.map((r) => `
  <section class="page" id="p${r.page}">
    <h2>Page ${r.page} <b>${percent(r.score)}</b> <small>pixel ${percent(r.pixel)} · layout ${percent(r.layout)}</small></h2>
    <div class="trio">
      <figure><figcaption>Design</figcaption><img src="${r.files.design}"></figure>
      <figure><figcaption>Preview</figcaption><img src="${r.files.preview ?? ""}" alt="missing"></figure>
      <figure><figcaption>Difference</figcaption><img src="${r.files.diff ?? ""}" alt="missing"></figure>
      <figure class="overlay"><figcaption>Overlay <input type="range" min="0" max="100" value="50"></figcaption>
        <div class="stack"><img src="${r.files.design}"><img class="top" src="${r.files.preview ?? ""}" style="opacity:.5"></div></figure>
    </div>
  </section>`).join("");

  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><title>${id} parity</title>
<style>
  :root { --ink:#16211c; --muted:#5c6b64; --line:#c8d3ce; --bg:#eef2f0; }
  body { margin:0; padding:16px; font:14px system-ui, sans-serif; color:var(--ink); background:var(--bg); }
  header { position:sticky; top:0; background:var(--bg); padding:8px 0; z-index:2; display:flex; gap:16px; align-items:baseline; flex-wrap:wrap; }
  header h1 { margin:0; font-size:20px; } header .score { font-size:28px; font-weight:700; }
  nav a { margin-right:8px; color:var(--muted); }
  .page { margin:24px 0; } .page h2 { font-size:16px; } .page h2 b { margin-left:8px; } .page small { color:var(--muted); font-weight:400; margin-left:8px; }
  .trio { display:grid; grid-template-columns:repeat(4, minmax(0, 1fr)); gap:12px; }
  figure { margin:0; } figcaption { font-size:12px; color:var(--muted); margin-bottom:4px; display:flex; gap:8px; align-items:center; }
  img { width:100%; display:block; background:#fff; box-shadow:0 0 0 1px var(--line); }
  .stack { position:relative; } .stack .top { position:absolute; inset:0; }
  @media (max-width:900px) { .trio { grid-template-columns:1fr 1fr; } }
</style></head><body>
<header><h1>${id}</h1><span class="score">${percent(overall)}</span><span>mean over ${results.length} pages</span>
<nav>${results.map((r) => `<a href="#p${r.page}">${r.page}: ${percent(r.score)}</a>`).join("")}</nav></header>
${rows}
<script>
for (const input of document.querySelectorAll(".overlay input")) {
  input.addEventListener("input", () => { input.closest(".overlay").querySelector(".top").style.opacity = input.value / 100; });
}
</script>
</body></html>`;
}

async function main() {
  const [id, ...rest] = process.argv.slice(2);
  if (!id) throw new Error("usage: design-parity.mjs <id> [--engine] [--pages 1,2]");
  const only = rest.includes("--pages") ? rest[rest.indexOf("--pages") + 1].split(",").map(Number) : null;
  // The preview build prints placeholders where the engine writes; the engine
  // build prints what it wrote. Each is compared with the canvas drawn the same
  // way, and neither with the canvas's on-screen highlight, which no file prints.
  const engine = rest.includes("--engine");
  // `--pdf <file>` scores a PDF — Word's export of the .docx — instead of the
  // HTML preview: the file itself, as the program that opens it draws it.
  const pdf = rest.includes("--pdf") ? resolve(rest[rest.indexOf("--pdf") + 1]) : null;
  const suffix = (engine ? ".engine" : "") + (pdf ? ".pdf" : "");

  const designDir = resolve(ROOT, "design", id);
  const canvas = (await readdir(designDir)).find((name) => name.endsWith(".dc.html") && !/Page\d/.test(name));
  const outDir = resolve(ROOT, ".parity", `${id}${suffix}`);
  await mkdir(outDir, { recursive: true });

  const server = await serve({ mode: engine ? "engine" : "preview", highlightAi: false });
  const origin = `http://127.0.0.1:${server.address().port}`;
  const page = await openChrome({ width: 1000, height: 1200 });

  try {
    const design = await photograph(page, `${origin}/design/${id}/${encodeURIComponent(canvas)}`, "article", 1);
    const previewName = `${id}${engine ? ".engine" : ""}.html`;
    const previewPath = resolve(ROOT, "public", "demo", previewName);
    const preview = pdf
      ? await photograph(page, `${origin}/__pdf?file=${encodeURIComponent("/" + pdf.slice(ROOT.length + 1))}`, "canvas.pdf-page", 1)
      : existsSync(previewPath)
      ? await photograph(page, `${origin}/public/demo/${previewName}`, "section.docx", 1)
      : { shots: [], texts: [] };
    const designs = design.shots;
    const previews = preview.shots;
    const drift = design.texts.map((texts, i) => ({ page: i + 1, ...textDrift(texts, preview.texts[i] ?? []) }));
    await writeFile(join(outDir, "text-drift.json"), JSON.stringify(drift, null, 2));

    const results = [];
    for (let i = 0; i < designs.length; i++) {
      const number = i + 1;
      if (only && !only.includes(number)) continue;
      const name = (kind) => `${kind}-${String(number).padStart(2, "0")}.png`;
      await writeFile(join(outDir, name("design")), PNG.sync.write(designs[i]));
      const files = { design: name("design") };
      let scores = { pixel: 0, layout: 0, score: 0 };
      if (previews[i]) {
        const { diff, ...rest } = compare(designs[i], previews[i]);
        scores = rest;
        await writeFile(join(outDir, name("preview")), PNG.sync.write(previews[i]));
        await writeFile(join(outDir, name("diff")), PNG.sync.write(diff));
        await writeFile(join(outDir, name("pair")), PNG.sync.write(pair(designs[i], previews[i])));
        files.preview = name("preview");
        files.diff = name("diff");
      }
      results.push({ page: number, ...scores, files });
    }

    await writeFile(join(outDir, "scores.json"), JSON.stringify({
      id,
      designPages: designs.length,
      previewPages: previews.length,
      overall: results.reduce((sum, r) => sum + r.score, 0) / results.length,
      pages: results.map(({ files, ...r }) => r),
    }, null, 2));
    await writeFile(join(outDir, "index.html"), board(id, results));

    console.log(`design ${designs.length} pages, preview ${previews.length} pages`);
    for (const r of results) {
      const d = drift[r.page - 1];
      console.log(`  page ${r.page}: ${percent(r.score)}  (pixel ${percent(r.pixel)}, layout ${percent(r.layout)})` +
        `  text: ${d.moved.length} off, ${d.missing.length} missing, ${d.extra.length} extra`);
    }
    console.log(`overall ${percent(results.reduce((sum, r) => sum + r.score, 0) / results.length)} → ${join(outDir, "index.html")}`);
  } finally {
    await page.close();
    server.close();
  }
}

await main();
