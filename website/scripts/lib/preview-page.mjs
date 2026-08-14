/**
 * Lifting the page out of the renderer's own output.
 *
 * renderDocumentWebsite emits a Word-alike shell: title bar, ribbon, ruler, and a
 * grey workspace around the page. That framing is right for the preview app and
 * wrong for a site, where the page — or a single node of it — is the subject.
 *
 * Hiding the shell with CSS overrides meant every embed depended on winning a
 * cascade fight with the preview app's grid; one stray rule and the page ended
 * up inset and clipped. Lifting the <style> and the <article> out instead means
 * there is no shell to fight: same stylesheet, same markup for the page, no
 * workspace, ruler or ribbon in the document at all.
 */

/** Whole-letter embeds: the A4 sheet, unframed, exactly as the renderer sized it. */
export const PAGE_ONLY_STYLE = `
      /* The page is the whole document here, so it owns the viewport. */
      html, body { margin: 0; padding: 0; background: transparent; }
      .a4-page { margin: 0; border: 0; box-shadow: none; }

      /* The "DYNAMIC: <id>" label is authoring information; here the page
         should read as the finished letter. */
      .node-id { display: none; }`;

/**
 * Single-node embeds: the same paper and the same typography, cropped to the
 * node. A4 proportions would put one paragraph in the top corner of an
 * otherwise empty sheet, which reads as an accident rather than an example.
 *
 * The DYNAMIC label stays: on a node page, whether a node announces itself as
 * unresolved is part of what the reader came to see.
 */
export const NODE_ONLY_STYLE = `
      html, body { margin: 0; padding: 0; background: transparent; }

      .a4-page {
        width: auto;
        min-height: 0;
        margin: 0;
        padding: 10mm 11mm;
        border: 0;
        box-shadow: none;
      }

      /* The letter's own title belongs to the letter, not to the node. */
      .document-title { display: none; }

      /* Margin above the first block and below the last would read as
         unexplained whitespace once the frame is sized to its content. The
         hidden title is still the first child, so the first VISIBLE block is
         addressed through it. */
      .a4-page > .document-title + * { margin-top: 0; }
      .document-section > h2:first-child { margin-top: 0; }
      .a4-page > :last-child { margin-bottom: 0; }
      .document-section > :last-child { margin-bottom: 0; }`;

/**
 * @param {string} html Output of renderDocumentWebsite.
 * @param {{ title: string, style: string }} options
 * @returns {string} A standalone document holding only the page.
 */
export function extractPage(html, { title, style: extraStyle }) {
  const styleStart = html.indexOf("<style>");
  const styleEnd = html.indexOf("</style>", styleStart);
  const pageStart = html.indexOf('<article class="a4-page"');
  const pageEnd = html.lastIndexOf("</article>");

  if (styleStart < 0 || styleEnd < 0 || pageStart < 0 || pageEnd < 0) {
    throw new Error(
      "Could not lift the page out of the renderer's output — its <style> or " +
        '<article class="a4-page"> is no longer where this script expects it.',
    );
  }

  const rendererStyle = html.slice(styleStart + "<style>".length, styleEnd);
  const page = html.slice(pageStart, pageEnd + "</article>".length);

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
    <style>${rendererStyle}</style>
    <style id="page-only">${extraStyle}
    </style>
  </head>
  <body>
${page}
  </body>
</html>
`;
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
