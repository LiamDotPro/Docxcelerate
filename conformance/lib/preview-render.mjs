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

import { JSDOM } from "jsdom";

/** The one window every case renders in. Building a fresh jsdom per case costs
 * about 300ms and buys nothing — docx-preview writes into the container it is
 * given, not into the document. */
let dom;

function browser() {
  if (dom === undefined) {
    dom = new JSDOM("<!doctype html><html><head></head><body></body></html>");

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
  }

  return dom.window;
}

/**
 * A document laid out by docx-preview, as the markup and CSS it produced.
 *
 * @param {object} model A finished document model.
 */
export async function renderPreview(model) {
  const window = browser();
  const [{ createDocxBlob }, { readPackedParagraphs, settleDocxPreview }, docxPreview] =
    await Promise.all([
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
  settleDocxPreview(bodyContainer, model, await readPackedParagraphs(packed));

  return { styles: styleContainer.innerHTML, pages: bodyContainer.innerHTML };
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
  const { styles, pages } = await renderPreview(model);

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
        document.body.dataset.paginated =
          JSON.stringify(DocxPaginate.paginateDocxPreview(document.body));
      } catch (error) {
        document.body.dataset.paginated = JSON.stringify({ error: String(error && error.message) });
      }
    </script>
  </body>
</html>
`;
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
    .replace(/docx_preview\.js$/, "docx_paginate.js");

  const bundled = await build({
    entryPoints: [entry],
    bundle: true,
    write: false,
    format: "iife",
    globalName: "DocxPaginate",
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
