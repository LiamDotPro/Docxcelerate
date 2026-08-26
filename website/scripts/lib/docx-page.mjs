/**
 * Rendering a document the only way that cannot drift: pack the `.docx` and
 * read it back.
 *
 * The site used to draw its previews with a second renderer that laid the same
 * model out in CSS. Two renderers means two answers, and the one nobody opened
 * in Word was free to be wrong — a footer bar whose text was navy on navy, a
 * total row in two different blues, tables that Word drew a black grid around
 * and the preview did not. None of that was in the document; it was in the
 * preview's opinion of the document.
 *
 * So the page here is produced the way the preview app in a scaffolded
 * workspace produces it: `createDocxBlob` writes the file the engine would
 * write, and `docx-preview` reads that file back and lays it out. What the
 * visitor sees is the file. If the packer changes, the site changes with it,
 * and there is no third place for a disagreement to hide.
 *
 * docx-preview is a browser library, so it is given a window to work in. That
 * is what jsdom is for: the pages are baked here, at build time, and the site
 * ships flat HTML rather than a renderer and a `.docx` to every visitor.
 */
import { settleDocxPreview } from "docxcelerate/preview";
import { settleDocxPreview } from "docxcelerate/preview";
import { JSDOM } from "jsdom";

/** The one window the whole build renders in. */
let dom;

/**
 * A DOM for a library that expects to be running in a page.
 *
 * The globals are set once and left: docx-preview reaches for `document` and
 * `DOMParser` directly rather than taking them as arguments, so they have to
 * be where a browser would keep them.
 */
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
    // The zip is read into blobs and the pictures are turned into data URIs
    // through a FileReader. Both have to come from the same window, or the
    // reader is handed a blob it does not recognise and every image in the
    // document comes out empty.
    globalThis.Blob = dom.window.Blob;
    globalThis.File = dom.window.File;
    globalThis.FileReader = dom.window.FileReader;
    globalThis.URL = dom.window.URL;
  }

  return dom.window;
}

/**
 * The pages of a document, as the HTML and CSS docx-preview lays them out in.
 *
 * @param {unknown} document A finished document model.
 * @returns {Promise<{ styles: string, pages: string }>}
 */
export async function renderDocxPreview(document) {
  const window = browser();
  const [{ createDocxBlob }, docxPreview] = await Promise.all([
    import("docxcelerate/docx"),
    import("docx-preview"),
  ]);

  const blob = await createDocxBlob(document);
  const styleContainer = window.document.createElement("div");
  const bodyContainer = window.document.createElement("div");

  await docxPreview.renderAsync(
    new Uint8Array(await blob.arrayBuffer()),
    bodyContainer,
    styleContainer,
    {
      // No grey desk around the sheet: the page is the whole subject here, and
      // the site frames it in its own chrome.
      inWrapper: false,
      // Every page break the document declares is a page break here.
      breakPages: true,
      ignoreLastRenderedPageBreak: false,
      // A standalone file cannot follow a blob: URL that died with the window
      // that made it, so pictures travel in the markup.
      useBase64URL: true,
    },
  );

  // What docx-preview did not read: a dropped field, a table indent it looks
  // for under the wrong attribute, a picture in an element a paragraph cannot
  // hold. The framework owns that now, so a scaffolded workspace and this site
  // finish a preview the same way rather than each discovering it separately.
  settleDocxPreview(bodyContainer, document);

  return {
    styles: withFontFallbacks(styleContainer.innerHTML),
    pages: withImageTypes(withFontFallbacks(bodyContainer.innerHTML)),
  };
}
/**
 * A font the reader does not have, falling back to one of the same kind.
 *
 * The file names its face and nothing else, which is right — it is a Word
 * document, and Word substitutes when a face is missing. A browser handed a
 * name it cannot resolve falls back to whatever the page inherits instead,
 * which is how a document set in a sans-serif comes out on screen in Times.
 * Naming the kind alongside the face keeps a substituted preview honest about
 * what the document is; it is a heuristic over the faces the shipped themes
 * use, not a font database.
 */
function withFontFallbacks(html) {
  return html.replaceAll(/font-family:\s*([^;"']+)/g, (rule, families) => {
    const listed = families.split(",").map((family) => family.trim());
    const generic = genericFor(listed[0]);

    return listed.includes(generic) ? rule : `font-family: ${families.trim()}, ${generic}`;
  });
}

/** Which kind of face a name belongs to. */
function genericFor(font) {
  const serif = ["cambria", "georgia", "times new roman", "garamond", "book antiqua", "palatino"];
  const mono = ["consolas", "courier new", "cascadia mono", "menlo"];
  const key = font.replaceAll(/["']/g, "").trim().toLowerCase();

  if (serif.includes(key)) {
    return "serif";
  }

  return mono.includes(key) ? "monospace" : "sans-serif";
}

/**
 * The pictures, labelled as pictures.
 *
 * docx-preview pulls each image out of the zip as a blob and JSZip hands it
 * back as `application/octet-stream`, which is what the data URI then says it
 * is. A browser believes that label and refuses to draw the picture, so a
 * document full of perfectly good images renders as a row of empty boxes. The
 * bytes are right; only the word for them is wrong, and the first few of them
 * say which word it should be.
 */
function withImageTypes(html) {
  return html.replaceAll(
    /data:application\/octet-stream;base64,([A-Za-z0-9+/=]+)/g,
    (uri, base64) => {
      const type = imageTypeOf(base64);

      return type === undefined ? uri : `data:${type};base64,${base64}`;
    },
  );
}

/** What a picture is, read from the first bytes of its base64. */
function imageTypeOf(base64) {
  const signatures = [
    ["iVBORw0KGgo", "image/png"],
    ["/9j/", "image/jpeg"],
    ["R0lGOD", "image/gif"],
    ["UklGR", "image/webp"],
    ["Qk", "image/bmp"],
    // An SVG is text, so it starts with whatever its first characters encode:
    // "<svg" or an XML declaration.
    ["PHN2Zy", "image/svg+xml"],
    ["PD94bWw", "image/svg+xml"],
  ];

  return signatures.find(([prefix]) => base64.startsWith(prefix))?.[1];
}

/**
 * One standalone page holding the rendered document and nothing else.
 *
 * @param {unknown} document A finished document model.
 * @param {{ title: string, style?: string }} options The page's title, and any
 * CSS the embed adds on top of docx-preview's own.
 * @returns {Promise<string>} The page, as HTML.
 */
export async function renderDocxPage(document, { title, style = PAGE_ONLY_STYLE }) {
  const { styles, pages } = await renderDocxPreview(document);

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- Embedded in an iframe by the page that explains it. On its own it is
         that page's content without its context, so it stays out of the index
         rather than competing with it. Crawling is still allowed, so the frame
         is not empty when the parent page is rendered. -->
    <meta name="robots" content="noindex">
    <title>${escapeHtml(title)}</title>
    ${styles}
    <style id="page-only">${style}
    </style>
  </head>
  <body>
${pages}
  </body>
</html>
`;
}

/**
 * Whole-document embeds: the sheet, unframed, at the size Word gives it.
 *
 * docx-preview writes the page's width, height and margins as inline styles
 * taken from the file, so nothing here touches the layout — only the ground it
 * sits on and the seam between one page and the next.
 */
export const PAGE_ONLY_STYLE = `
      html, body { margin: 0; padding: 0; background: transparent; }

      section.docx {
        margin: 0 auto;
        background: #ffffff;
        box-sizing: border-box;
      }

      section.docx + section.docx { margin-top: 26px; }`;

/**
 * Single-node embeds: the same paper and the same typography, cropped to the
 * node. A4 proportions would put one paragraph in the top corner of an
 * otherwise empty sheet, which reads as an accident rather than an example.
 */
export const NODE_ONLY_STYLE = `
      html, body { margin: 0; padding: 0; background: transparent; }

      section.docx {
        width: auto !important;
        min-height: 0 !important;
        margin: 0;
        padding: 10mm 11mm !important;
        background: transparent;
      }

      /* Margin above the first block and below the last would read as
         unexplained whitespace once the frame is sized to its content. */
      section.docx > :first-child { margin-top: 0; }
      section.docx > :last-child { margin-bottom: 0; }`;

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
