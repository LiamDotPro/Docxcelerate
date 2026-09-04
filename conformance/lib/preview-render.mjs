/**
 * The preview, rendered the way the framework says a preview is rendered:
 * pack the `.docx`, and read it back with docx-preview.
 *
 * There is no second renderer here and there must never be one. The whole
 * reason this probe can be evidence is that it lays out the same bytes Word
 * opens — a page drawn from the model in CSS would be a second opinion, and
 * the one nobody opens in Word is the one free to be wrong.
 *
 * docx-preview is a browser library, so it is given a window: jsdom does not
 * lay anything out, but it is enough to produce the markup and CSS that
 * headless Chrome then measures for real.
 *
 * @module
 */

import { JSDOM, VirtualConsole } from "jsdom";
import { mkdir } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");

/** The one window every case renders in. Building a fresh jsdom per case costs
 * about 300ms and buys nothing — docx-preview writes into the container it is
 * given, not into the document. */
let dom;

function browser() {
  if (dom === undefined) {
    // jsdom shouts about the one thing it cannot do that is asked of it here:
    // ECharts measures text through a canvas, jsdom has no `getContext`, and
    // ECharts falls back to estimating the width. That is fine — the frame is
    // what this tier measures and the frame comes from the file — but one
    // "Not implemented" per label buries the board it is printed above.
    const quiet = new VirtualConsole();
    quiet.on("jsdomError", (error) => {
      if (!String(error?.message ?? "").includes("getContext")) {
        console.error(error);
      }
    });

    dom = new JSDOM("<!doctype html><html><head></head><body></body></html>", {
      virtualConsole: quiet,
    });

    globalThis.window = dom.window;
    globalThis.document = dom.window.document;
    globalThis.DOMParser = dom.window.DOMParser;
    globalThis.XMLSerializer = dom.window.XMLSerializer;
    globalThis.Node = dom.window.Node;
    globalThis.Element = dom.window.Element;
    globalThis.HTMLElement = dom.window.HTMLElement;
    globalThis.getComputedStyle = dom.window.getComputedStyle;
    // The zip is read into blobs and pictures become data URIs through a
    // FileReader; both have to come from the same window, or the reader is
    // handed a blob it does not recognise and every image comes out empty.
    globalThis.Blob = dom.window.Blob;
    globalThis.File = dom.window.File;
    globalThis.FileReader = dom.window.FileReader;
    globalThis.URL = dom.window.URL;
    // docx-preview sizes a VML shape's <svg> from its rendered bounding box,
    // one animation frame after it draws it. jsdom has neither the frame nor
    // the box, so without a stand-in the render throws and a document with a
    // shape in it bakes to nothing. A no-op is the right stand-in rather than
    // jsdom's own `pretendToBeVisual` clock: the callback would then run and
    // die on the getBBox jsdom does not implement, and the number it computes
    // is one the <svg> already carries in the style the file gave it.
    globalThis.requestAnimationFrame = () => 0;
    globalThis.cancelAnimationFrame = () => {};
  }

  return dom.window;
}

/**
 * The workspace's own chart drawer, bundled once and kept.
 *
 * It is a TypeScript file under `templates/`, so esbuild turns it into
 * something Node can import — the same recipe `loadCase` uses for a case, with
 * the packages left external so `echarts` resolves out of this package's own
 * `node_modules`.
 */
let drawer;

async function chartDrawer() {
  if (drawer === undefined) {
    const { build } = await import("esbuild");
    const source = resolve(ROOT, "..", "templates", "workspace", "preview", "charts.ts");
    const bundle = resolve(ROOT, ".out", "_bundles", "chart-drawer.mjs");

    await mkdir(dirname(bundle), { recursive: true });
    await build({
      entryPoints: [source],
      outfile: bundle,
      bundle: true,
      format: "esm",
      platform: "node",
      target: "node20",
      packages: "external",
      logLevel: "warning",
    });

    const module = await import(pathToFileURL(bundle).href);
    drawer = module.createChartDrawer();
  }

  return drawer;
}

/**
 * A document laid out by docx-preview, as the markup and CSS it produced.
 *
 * @param {object} model A finished document model.
 */
export async function renderPreview(model) {
  const window = browser();
  const [
    { createDocxBlob },
    {
      readPackedCharts,
      readPackedParagraphs,
      readPackedTables,
      settleDocxPreview,
      settleDocxPreviewCharts,
    },
    docxPreview,
  ] = await Promise.all([
    import("docxcelerate/docx"),
    import("docxcelerate/preview"),
    import("docx-preview"),
  ]);

  const blob = await createDocxBlob(model);
  const packed = new Uint8Array(await blob.arrayBuffer());
  const styleContainer = window.document.createElement("div");
  const bodyContainer = window.document.createElement("div");

  await docxPreview.renderAsync(packed, bodyContainer, styleContainer, {
    inWrapper: false,
    breakPages: true,
    ignoreLastRenderedPageBreak: false,
    useBase64URL: true,
  });

  // What docx-preview did not read: a dropped field, a table indent under the
  // wrong attribute, a run's letter spacing, a border's gap to its text, a
  // picture in an element a paragraph cannot hold. The framework owns finishing
  // that, and a conformance run has to measure the preview a person actually
  // sees — which is the settled one, read back from the very bytes it drew.
  settleDocxPreview(
    bodyContainer,
    model,
    await readPackedParagraphs(packed),
    await readPackedTables(packed),
  );

  // The charts, drawn by the same drawer a scaffolded workspace uses.
  //
  // docx-preview leaves a chart as an empty frame of the right size, so
  // without this the preview tier would measure a hole and a case about charts
  // could never be anything but red. The drawer is not the suite's own — it is
  // `templates/workspace/preview/charts.ts`, bundled straight out of the
  // template directory, so what is measured here is what a person previewing a
  // document actually sees. Writing a second one for the harness would make
  // the preview tier evidence about the harness.
  settleDocxPreviewCharts(bodyContainer, await readPackedCharts(packed), await chartDrawer());

  return {
    styles: styleContainer.innerHTML,
    pages: bodyContainer.innerHTML,
    running: await renderRunningFurniture(packed, docxPreview, window),
    packed: await readPackedParagraphs(packed),
  };
}

/**
 * The strip a title-page document draws on every page *after* the first.
 *
 * docx-preview picks one header and one footer per page it renders, and it
 * renders one page — so for a document whose first page differs it draws the
 * letterhead and never asks for the running header at all. The default part is
 * in the package and Word draws it from page two; it simply never reaches the
 * DOM, so the paginator has nothing but the letterhead to carry forward.
 *
 * The fix is to ask docx-preview for it, rather than to build it. The file is
 * parsed, the section's title-page flag is turned off in the parsed tree only,
 * and the same renderer draws the same document again into a scratch container
 * — where page one now takes the *default* parts. What comes out is
 * docx-preview's rendering of the document's own running header, which is the
 * only thing that would not be a second opinion.
 *
 * Nothing is written back to the file, and the real preview is rendered from
 * the untouched bytes. A document that is not a title page skips this entirely.
 */
async function renderRunningFurniture(packed, docxPreview, window) {
  const parsed = await docxPreview.parseAsync(packed);
  // The section properties live on the body itself — `body.props` — not on a
  // `sections` list. A document with one section has one set of them, and it is
  // the same object `renderHeaderFooter` consults.
  const properties = parsed?.documentPart?.body?.props;

  if (properties === undefined) {
    return { header: null, footer: null, evenHeader: null, evenFooter: null };
  }

  const running = properties.titlePage === true
    ? await drawStrips(parsed, docxPreview, window, () => {
      properties.titlePage = false;
    })
    : { header: null, footer: null };

  // The verso's strip, by the same means. docx-preview chooses `even` only for
  // a page it is rendering second, and it renders one — so instead the *even*
  // reference is put where the default one goes and the same renderer is asked
  // again. It draws the document's own even-page part; nothing about the file
  // changes, and the real preview is rendered from the untouched bytes.
  const hasEven = (properties.headerRefs ?? []).some((ref) => ref.type === "even") ||
    (properties.footerRefs ?? []).some((ref) => ref.type === "even");

  const even = hasEven
    ? await drawStrips(parsed, docxPreview, window, () => {
      properties.titlePage = false;
      promoteEven(properties.headerRefs);
      promoteEven(properties.footerRefs);
    })
    : { header: null, footer: null };

  return {
    header: running.header,
    footer: running.footer,
    evenHeader: even.header,
    evenFooter: even.footer,
  };
}

/** The even reference, moved to where docx-preview looks for the default one. */
function promoteEven(refs) {
  const even = (refs ?? []).find((ref) => ref.type === "even");
  const fallback = (refs ?? []).find((ref) => ref.type === "default");

  if (even !== undefined && fallback !== undefined) {
    fallback.id = even.id;
  }
}

/** One scratch render, and the strips it drew. */
async function drawStrips(parsed, docxPreview, window, prepare) {
  prepare();

  const body = window.document.createElement("div");
  const head = window.document.createElement("div");

  await docxPreview.renderDocument(parsed, body, head, {
    inWrapper: false,
    breakPages: true,
    ignoreLastRenderedPageBreak: false,
    useBase64URL: true,
  });

  const page = body.querySelector("section.docx");

  return {
    header: page?.querySelector(":scope > header")?.outerHTML ?? null,
    footer: page?.querySelector(":scope > footer")?.outerHTML ?? null,
  };
}

/**
 * One standalone page holding the rendered document and nothing else.
 *
 * The only CSS added on top of docx-preview's own paints the paper and puts a
 * seam between pages. Nothing here touches type, spacing or position: this
 * page is the thing under test, and a stylesheet that corrected it would be
 * the harness marking its own homework.
 */
export async function renderPreviewPage(model, title) {
  const { styles, pages, running, packed } = await renderPreview(model);

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <title>${escapeHtml(title ?? "case")}</title>
    ${styles}
    <style id="paper">
      html, body { margin: 0; padding: 0; background: #d8dade; }
      section.docx { margin: 0 auto; background: #ffffff; box-sizing: border-box; }
      section.docx + section.docx { margin-top: 26px; }
    </style>
  </head>
  <body>
${pages}
${runningMarkup(running)}
    <script id="packed" type="application/json">${escapeJson(packed)}</script>
    <script>
${await paginatorSource()}
      // Pagination is the one step that cannot happen at bake time: it needs to
      // know how tall things actually drew, and jsdom lays nothing out. So it
      // runs here, in the browser about to measure the page — which is where a
      // reader's browser would run it too.
      //
      // The paginator is inlined rather than imported. An external module
      // fetched from inside an iframe never loads under Chrome's
      // --virtual-time-budget: the clock runs out before the request is served,
      // and the page comes back unpaginated with no error to say why. Inline,
      // there is nothing to fetch. The source is read from the built package,
      // so this is the framework's own paginator and not a copy of it.
      try {
        // Tab stops before pagination: a stop changes how tall a paragraph is
        // only by keeping it on one line, and a line that has not been placed
        // yet is the wrong height to paginate against.
        const packed = JSON.parse(document.getElementById("packed").textContent);
        document.body.dataset.tabs = String(DocxPreviewKit.applyTabStops(document.body, packed));

        document.body.dataset.paginated = JSON.stringify(
          DocxPreviewKit.paginateDocxPreview(document.body, {
            runningHeader: document.getElementById("running-header")?.firstElementChild ?? null,
            runningFooter: document.getElementById("running-footer")?.firstElementChild ?? null,
            evenHeader: document.getElementById("even-header")?.firstElementChild ?? null,
            evenFooter: document.getElementById("even-footer")?.firstElementChild ?? null,
          }),
        );
      } catch (error) {
        document.body.dataset.paginated = JSON.stringify({ error: String(error && error.message) });
      }
    </script>
  </body>
</html>
`;
}

/**
 * What the file says about its paragraphs, safe to sit in a `<script>`.
 *
 * Only `<` needs escaping, and only because a `</script` anywhere inside the
 * block would end it early — a paragraph whose text happens to say so would
 * otherwise take the rest of the page with it.
 */
function escapeJson(value) {
  return JSON.stringify(value).replaceAll("<", "\\u003c");
}

/**
 * The running furniture, parked in the page for the paginator to find.
 *
 * Hidden, and outside every sheet: it is not drawn, it is the pattern each new
 * sheet is given after the first. `display: none` would cost it its layout and
 * with it the reserve heights the paginator reads, so it is moved off the page
 * instead.
 */
function runningMarkup(running) {
  if (!running.header && !running.footer && !running.evenHeader && !running.evenFooter) {
    return "";
  }

  const parked = 'style="position:absolute;left:-99999px;top:0"';

  const strips = [
    ["running-header", running.header],
    ["running-footer", running.footer],
    ["even-header", running.evenHeader],
    ["even-footer", running.evenFooter],
  ];

  return strips
    .filter(([, markup]) => Boolean(markup))
    .map(([id, markup]) => `    <div id="${id}" ${parked}>${markup}</div>`)
    .join("\n");
}

/**
 * The built paginator, bundled into one self-contained script.
 *
 * Bundled rather than read and string-edited. The first version stripped the
 * `export` keywords and inlined the file, which worked right up to the moment
 * the paginator grew an `import` of its own — the browser then tried to fetch a
 * module that was never served, the whole inline script failed, and the page
 * came back unpaginated with nothing to say why. Every case still passed,
 * because each was self-consistently measuring a one-page document.
 *
 * esbuild resolves the imports for real, so this cannot rot the same way again.
 */
async function paginatorSource() {
  const { build } = await import("esbuild");
  const { createRequire } = await import("node:module");
  const entry = createRequire(import.meta.url)
    .resolve("docxcelerate/preview")
;

  const bundled = await build({
    entryPoints: [entry],
    bundle: true,
    write: false,
    format: "iife",
    globalName: "DocxPreviewKit",
    target: "es2022",
    logLevel: "error",
  });

  return bundled.outputFiles[0].text;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
