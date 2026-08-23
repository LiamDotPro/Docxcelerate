import type {
  DocumentBlockStyle,
  DocumentModel,
  DocumentNode,
  DocumentStyle,
  TableNode,
  TableRowNode,
} from "../domain/types.ts";
import { cleanMinimalDocumentStyle } from "../project/style.ts";
import { imageSourceOf } from "../render/image_source.ts";

/**
 * Rendering a document to HTML, for reading it on a screen.
 *
 * @module
 */

/** What {@linkcode renderDocumentWebsite} takes beyond the document. */
export interface WebRenderOptions {
  /** Reload the page when the document changes, for the preview app. */
  liveReload?: boolean;
  /** The page title. Defaults to the document's own. */
  title?: string;
}

/**
 * Renders a document as a standalone HTML page, styles inlined and nothing
 * fetched.
 *
 * @param doc The document to render.
 * @param options The page title, and whether to poll for changes.
 * @returns The page, as HTML.
 */
export function renderDocumentWebsite(
  doc: DocumentModel,
  options: WebRenderOptions = {},
): string {
  const title = options.title ?? doc.title;
  // The style travels with the document, so a preview shows the theme the
  // project chose rather than the one this renderer was written against.
  const style = doc.style ?? cleanMinimalDocumentStyle;

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>${escapeHtml(title)}</title>
    <style>
      :root {
        color-scheme: light;
        font-family: "Aptos", "Segoe UI", Arial, sans-serif;
        background: #f3f4f7;
        color: #1f2933;
        --chrome: #fbfbfd;
        --chrome-border: #d8dce3;
        --workspace: #e6e8ed;
        --page-border: #c7ccd6;
        --text-muted: #657287;
${styleVariables(style)}
      }

      * {
        box-sizing: border-box;
      }

      html,
      body {
        margin: 0;
        min-height: 100vh;
      }

      body {
        overflow: hidden;
        background: var(--workspace);
      }

      .word-shell {
        display: grid;
        grid-template-rows: auto auto 1fr;
        min-height: 100vh;
      }

      .titlebar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        min-height: 44px;
        padding: 0 18px;
        background: #204c91;
        color: #ffffff;
      }

      .titlebar strong {
        font-size: 14px;
        font-weight: 600;
      }

      .titlebar-meta {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        font-size: 12px;
        opacity: .94;
      }

      .titlebar-meta::before {
        content: "";
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background: ${options.liveReload ? "#44c767" : "#a9bad8"};
      }

      .ribbon {
        background: var(--chrome);
        border-bottom: 1px solid var(--chrome-border);
        box-shadow: 0 1px 0 rgb(15 23 42 / 4%);
      }

      .tabs {
        display: flex;
        gap: 2px;
        min-height: 34px;
        padding: 0 16px;
        border-bottom: 1px solid var(--chrome-border);
      }

      .tab {
        display: inline-flex;
        align-items: center;
        padding: 0 14px;
        color: #384252;
        font-size: 13px;
      }

      .tab-active {
        border-bottom: 2px solid var(--accent);
        color: var(--accent);
        font-weight: 600;
      }

      .tools {
        display: flex;
        align-items: center;
        gap: 22px;
        min-height: 74px;
        padding: 10px 18px 12px;
      }

      .tool-group {
        display: grid;
        grid-template-columns: repeat(3, 30px);
        gap: 6px;
        padding-right: 20px;
        border-right: 1px solid var(--chrome-border);
      }

      .tool {
        height: 26px;
        border: 1px solid #d4d9e2;
        border-radius: 4px;
        background: linear-gradient(#ffffff, #f0f2f6);
      }

      .tool-wide {
        width: 116px;
        height: 26px;
        border: 1px solid #d4d9e2;
        border-radius: 4px;
        background: #ffffff;
      }

      .workspace {
        overflow: auto;
        background:
          linear-gradient(90deg, rgb(0 0 0 / 4%) 0 1px, transparent 1px) 0 0 / 24px 24px,
          var(--workspace);
      }

      .workspace-inner {
        width: max-content;
        min-width: 1180px;
        min-height: calc(100vh - 153px);
        padding: 0 96px 80px;
      }

      .horizontal-ruler {
        position: sticky;
        top: 0;
        z-index: 1;
        width: var(--page-width);
        height: 28px;
        margin: 0 auto 18px;
        background:
          repeating-linear-gradient(
            90deg,
            #f8f9fb 0,
            #f8f9fb 9.5mm,
            #c7ccd6 9.5mm,
            #c7ccd6 10mm
          );
        border: 1px solid #cfd4dd;
        border-top: 0;
        box-shadow: 0 1px 2px rgb(15 23 42 / 10%);
      }

      .page-stage {
        display: grid;
        place-items: start center;
        min-width: 100%;
      }

      .a4-page {
        display: flex;
        flex-direction: column;
        width: var(--page-width);
        min-height: var(--page-height);
        background: var(--page);
        border: 1px solid var(--page-border);
        box-shadow:
          0 1px 2px rgb(15 23 42 / 12%),
          0 16px 44px rgb(15 23 42 / 24%);
      }

      .a4-page + .a4-page {
        margin-top: 26px;
      }

      /* The body carries the margins; the running furniture sits outside them,
         which is what a margin is for. */
      .page-body {
        flex: 1;
        padding: var(--page-padding);
      }

      .page-header {
        padding-top: 8mm;
      }

      .page-footer {
        margin-top: auto;
      }

      /* Furniture sits in the margins horizontally, but a block that is meant
         to be a line across the page runs the whole width of it. */
      .page-header > *,
      .page-footer > * {
        margin: 0;
        padding-right: var(--page-margin-right);
        padding-left: var(--page-margin-left);
      }

      .page-header > [data-variant],
      .page-footer > [data-variant] {
        padding-right: 0;
        padding-left: 0;
      }

      .page-header .doc-paragraph,
      .page-footer .doc-paragraph {
        margin: 0;
      }

      .page-header .doc-table td,
      .page-footer .doc-table td {
        border-bottom: 0;
      }

      .page-number[data-format="currentOfTotal"]::before {
        content: var(--page-current) attr(data-separator) var(--page-total);
      }

      .page-number[data-format="current"]::before {
        content: var(--page-current);
      }

      .page-number[data-format="total"]::before {
        content: var(--page-total);
      }

      .document-title {
        margin: var(--title-space-before) 0 var(--title-space-after);
        font-family: var(--heading-font);
        font-size: var(--title-size);
        line-height: 1.2;
        font-weight: var(--title-weight);
        text-transform: var(--title-transform);
        color: var(--title-color);
      }

      .doc-section {
        margin: 0 0 12pt;
      }

      .doc-section h2 {
        margin: var(--heading-space-before) 0 var(--heading-space-after);
        font-family: var(--heading-font);
        font-size: var(--heading-size);
        line-height: 1.3;
        font-weight: var(--heading-weight);
        text-transform: var(--heading-transform);
        color: var(--heading-color);
      }

      .doc-paragraph {
        margin: 0 0 var(--paragraph-space-after);
        font-family: var(--body-font);
        font-size: var(--body-size);
        line-height: var(--body-line-height);
        color: var(--body-color);
      }

      .doc-table {
        width: 100%;
        margin: 0 0 var(--paragraph-space-after);
        border-collapse: collapse;
        table-layout: fixed;
        font-family: var(--body-font);
        font-size: var(--body-size);
        color: var(--body-color);
      }

      .doc-table th {
        padding: 6pt 8pt;
        background: var(--accent);
        color: var(--page);
        font-size: .72em;
        font-weight: 600;
        letter-spacing: .1em;
        text-transform: uppercase;
        vertical-align: bottom;
      }

      .doc-table td {
        padding: 5pt 8pt;
        border-bottom: 1px solid var(--rule);
        vertical-align: baseline;
        overflow-wrap: break-word;
      }

      .doc-table tbody tr:nth-child(even) td {
        background: color-mix(in srgb, var(--accent) 4%, var(--page));
      }

      .doc-table .doc-paragraph {
        margin: 0;
      }

      .doc-table .doc-paragraph + .doc-paragraph {
        margin-top: 1pt;
        color: var(--muted);
        font-size: .84em;
      }

      .node-id {
        display: block;
        margin-bottom: 3pt;
        color: var(--text-muted);
        font-size: 8pt;
        line-height: 1.2;
        text-transform: uppercase;
      }

      .doc-image {
        display: block;
        max-width: 100%;
      }

      .image-placeholder,
      .graph-placeholder,
      .toc-placeholder {
        margin: 10pt 0;
        padding: 9pt;
        border: 1px dashed var(--rule);
        color: var(--muted);
        font-size: 9pt;
      }

${blockCss(style)}
      @media (max-width: 900px) {
        body {
          overflow: auto;
        }

        .word-shell {
          min-width: 900px;
        }
      }
    </style>
  </head>
  <body>
    <main class="word-shell">
      <header class="titlebar">
        <strong>${escapeHtml(doc.title)}</strong>
        <span class="titlebar-meta">${
    options.liveReload ? "Live preview connected" : "Preview"
  }</span>
      </header>

      <section class="ribbon" aria-label="Document toolbar">
        <nav class="tabs" aria-label="Ribbon tabs">
          <span class="tab tab-active">Home</span>
          <span class="tab">Insert</span>
          <span class="tab">Layout</span>
          <span class="tab">Review</span>
          <span class="tab">View</span>
        </nav>
        <div class="tools" aria-hidden="true">
          <div class="tool-wide"></div>
          <div class="tool-group">
            <span class="tool"></span>
            <span class="tool"></span>
            <span class="tool"></span>
            <span class="tool"></span>
            <span class="tool"></span>
            <span class="tool"></span>
          </div>
          <div class="tool-wide"></div>
          <div class="tool-wide"></div>
        </div>
      </section>

      <section class="workspace" aria-label="Document workspace">
        <div class="workspace-inner">
          <div class="horizontal-ruler" aria-hidden="true"></div>
          <div class="page-stage">
            ${renderPages(doc, style)}
          </div>
        </div>
      </section>
    </main>
    ${options.liveReload ? liveReloadScript() : ""}
  </body>
</html>`;
}

/**
 * The pages, split where the document said to break.
 *
 * A preview that runs everything into one sheet is not showing the document —
 * an invoice whose payment details belong on their own page has to be seen to
 * put them there. The break is the document's; how much else fits is the
 * browser's business, and it is not asked.
 *
 * The running header and footer are drawn on each page, so what repeats looks
 * like it repeats. The title goes on the first page only: it is the document's
 * name, not furniture.
 */
function renderPages(doc: DocumentModel, style: DocumentStyle): string {
  const pages = splitPages(doc.nodes);
  const showTitle = style.showTitle !== false;
  const header = doc.header?.map(renderNode).join("\n") ?? "";
  const footer = doc.footer?.map(renderNode).join("\n") ?? "";

  return pages
    .map((nodes, index) => {
      const label = pages.length === 1
        ? escapeHtml(doc.title)
        : `${escapeHtml(doc.title)}, page ${index + 1} of ${pages.length}`;

      return `<article class="a4-page" aria-label="${label}" data-page="${index + 1}"
        style="--page-current:'${index + 1}';--page-total:'${pages.length}'">
        ${header === "" ? "" : `<div class="page-header">${header}</div>`}
        <div class="page-body">
          ${index === 0 && showTitle ? `<h1 class="document-title">${escapeHtml(doc.title)}</h1>` : ""}
          ${nodes.map(renderNode).join("\n")}
        </div>
        ${footer === "" ? "" : `<div class="page-footer">${footer}</div>`}
      </article>`;
    })
    .join("\n");
}

/** Splits a run of nodes at every page break, dropping the breaks themselves. */
function splitPages(nodes: readonly DocumentNode[]): DocumentNode[][] {
  const pages: DocumentNode[][] = [[]];

  for (const node of nodes) {
    if (node.kind === "pageBreak") {
      pages.push([]);
      continue;
    }

    pages[pages.length - 1].push(node);
  }

  // A break at the very end asks for a page nothing goes on. The document is
  // what it is, but an empty sheet in a preview reads as a bug rather than as
  // a decision, so it is dropped.
  return pages.filter((page, index) => page.length > 0 || index === 0);
}

function renderNode(node: DocumentNode): string {
  if (node.kind === "pageBreak") {
    return "";
  }

  if (node.kind === "pageNumber") {
    return `<span class="page-number" data-node-id="${escapeHtml(node.id)}"
      data-format="${escapeHtml(node.format ?? "currentOfTotal")}"
      data-separator="${escapeHtml(node.separator ?? " / ")}"></span>`;
  }

  if (node.kind === "section") {
    return `<section class="doc-section" data-node-id="${escapeHtml(node.id)}"${variantAttr(node)}>
      <h2>${escapeHtml(node.title ?? node.id)}</h2>
      ${node.children.map(renderNode).join("\n")}
    </section>`;
  }

  if (node.kind === "paragraph") {
    const className = node.mode === "dynamic"
      ? "doc-paragraph paragraph-dynamic"
      : "doc-paragraph paragraph-static";

    return `<p class="${className}" data-node-id="${escapeHtml(node.id)}"${variantAttr(node)}>
      ${
      node.mode === "dynamic" ? `<span class="node-id">Dynamic: ${escapeHtml(node.id)}</span>` : ""
    }
      ${escapeHtml(node.text ?? "")}
    </p>`;
  }

  if (node.kind === "image") {
    const source = imageSourceOf(node.path);

    // A picture the page can actually show is shown. The dashed box is for a
    // node that has no picture yet — one the engine will produce — and saying
    // so is the point of it; drawing a box where an image exists would be the
    // preview lying about what the document contains.
    if (source.kind !== "none") {
      const size = [
        node.width === undefined ? "" : `width:${node.width}pt;`,
        node.height === undefined ? "" : `height:${node.height}pt;`,
      ].join("");

      return `<img class="doc-image" data-node-id="${escapeHtml(node.id)}"${variantAttr(node)}
        src="${escapeHtml(source.uri)}" alt="${escapeHtml(node.alt ?? "")}"
        ${size === "" ? "" : `style="${size}"`}>`;
    }

    return `<div class="image-placeholder" data-node-id="${escapeHtml(node.id)}"${
      variantAttr(node)
    }>
      Image: ${escapeHtml(node.alt ?? node.placeholder ?? node.id)}
    </div>`;
  }

  if (node.kind === "graph") {
    return `<div class="graph-placeholder" data-node-id="${escapeHtml(node.id)}"${variantAttr(node)}>
      ${escapeHtml(node.graphType)} graph: ${
      escapeHtml(node.caption ?? node.placeholder ?? node.id)
    }
    </div>`;
  }

  if (node.kind === "repeat") {
    return `<div class="repeat-body" data-node-id="${escapeHtml(node.id)}">
      <span class="node-id">Repeats per ${escapeHtml(node.source.path)}</span>
      ${node.children.map(renderNode).join("\n")}
    </div>`;
  }

  if (node.kind === "table") {
    return renderTable(node);
  }

  // A row or a cell outside the table it belongs to. Drawing the children is
  // the honest thing to do — it is a build that put them there, and dropping
  // them would hide the mistake rather than show it.
  if (node.kind === "tableRow" || node.kind === "tableCell") {
    return node.children.map(renderNode).join("\n");
  }

  return `<div class="toc-placeholder" data-node-id="${escapeHtml(node.id)}"${variantAttr(node)}>
    ${escapeHtml(node.title ?? "Table of contents")}
  </div>`;
}

/**
 * The document's style, as custom properties for the page CSS to read.
 *
 * Everything the stylesheet needs is declared here rather than interpolated
 * into forty rules, so the rules below stay readable as CSS and a theme is one
 * block to look at when a preview comes out wrong.
 */
function styleVariables(style: DocumentStyle): string {
  const { page, typography, paragraph, title, sectionHeading } = style;
  const landscape = page.orientation === "landscape";
  const shortSideMm = page.size === "A4" ? 210 : 215.9;
  const longSideMm = page.size === "A4" ? 297 : 279.4;

  return [
    ["--page-width", `${landscape ? longSideMm : shortSideMm}mm`],
    ["--page-height", `${landscape ? shortSideMm : longSideMm}mm`],
    [
      "--page-padding",
      `${page.margins.topMm}mm ${page.margins.rightMm}mm ` +
      `${page.margins.bottomMm}mm ${page.margins.leftMm}mm`,
    ],
    ["--page", cssColor(style.palette?.page ?? "FFFFFF")],
    ["--page-margin-left", `${page.margins.leftMm}mm`],
    ["--page-margin-right", `${page.margins.rightMm}mm`],
    ["--body-font", fontStack(typography.bodyFont)],
    ["--heading-font", fontStack(typography.headingFont)],
    ["--body-size", `${typography.bodySizePt}pt`],
    ["--body-line-height", String(typography.bodyLineHeight)],
    ["--body-color", cssColor(typography.color)],
    ["--paragraph-space-after", `${paragraph.spacingAfterPt}pt`],
    ["--title-size", `${title.fontSizePt}pt`],
    ["--title-weight", title.weight === "bold" ? "700" : "400"],
    ["--title-color", cssColor(title.color ?? style.palette?.heading ?? typography.color)],
    ["--title-space-before", `${title.spacingBeforePt}pt`],
    ["--title-space-after", `${title.spacingAfterPt}pt`],
    ["--title-transform", title.transform === "uppercase" ? "uppercase" : "none"],
    ["--heading-size", `${sectionHeading.fontSizePt}pt`],
    ["--heading-weight", sectionHeading.weight === "bold" ? "700" : "400"],
    [
      "--heading-color",
      cssColor(sectionHeading.color ?? style.palette?.heading ?? typography.color),
    ],
    ["--heading-space-before", `${sectionHeading.spacingBeforePt}pt`],
    ["--heading-space-after", `${sectionHeading.spacingAfterPt}pt`],
    ["--heading-transform", sectionHeading.transform === "uppercase" ? "uppercase" : "none"],
    ["--accent", cssColor(style.palette?.accent ?? typography.color)],
    ["--muted", cssColor(style.palette?.muted ?? "59677A")],
    ["--rule", cssColor(style.palette?.rule ?? "9AA6B8")],
  ]
    .map(([name, value]) => `        ${name}: ${value};`)
    .join("\n");
}

/**
 * A font name as a CSS stack, with a fallback in the same key.
 *
 * Word substitutes for a missing font silently and a browser falls back to
 * whatever the page inherits, which is how a serif theme ends up previewing in
 * a sans-serif face on a machine without the font. Naming the family the theme
 * belongs to keeps the preview approximately honest when the exact face is not
 * installed; it is a heuristic over the fonts the shipped themes use, not a
 * font database.
 */
function fontStack(font: string): string {
  const serif = ["cambria", "georgia", "times new roman", "garamond", "book antiqua", "palatino"];
  const mono = ["consolas", "courier new", "cascadia mono", "menlo"];
  const key = font.trim().toLowerCase();
  const generic = serif.includes(key) ? "serif" : mono.includes(key) ? "monospace" : "sans-serif";

  return `"${cssIdentifier(font)}", ${generic}`;
}

/** A style colour as CSS. Stored without the hash, because OOXML wants it that way. */
function cssColor(value: string): string {
  return /^[0-9a-fA-F]{6}$/.test(value) ? `#${value}` : cssIdentifier(value);
}

/**
 * Anything from the style that lands inside the `<style>` block, with the
 * characters that could leave it removed.
 *
 * The style is typed and usually written by hand, but it also arrives inside a
 * published document — which is data, and data from a request has no business
 * closing a style tag.
 */
function cssIdentifier(value: string): string {
  return value.replace(/[^a-zA-Z0-9 \-_.#%(),]/g, "");
}

/**
 * Draws a table, with the header rows in a `<thead>` so they repeat on print.
 *
 * A published table carries loops rather than rows, because how many lines an
 * invoice has belongs to the request. Those are walked into the rows they
 * stand for, so a preview shows the table the engine will write rather than a
 * note saying there would be one.
 */
function renderTable(node: TableNode): string {
  const rows = tableRows(node.children);
  const lead = leadingHeaders(rows);
  const head = rows.slice(0, lead);
  const body = rows.slice(lead);
  const cols = node.columns
    .map((column) =>
      `<col${
        column.width === undefined || column.width === "auto"
          ? ""
          : ` style="width:${column.width}mm"`
      }>`
    )
    .join("");

  return `<table class="doc-table" data-node-id="${escapeHtml(node.id)}"${variantAttr(node)}>
    <colgroup>${cols}</colgroup>
    ${head.length > 0 ? `<thead>${head.map((row) => renderRow(row, node)).join("\n")}</thead>` : ""}
    <tbody>${body.map((row) => renderRow(row, node)).join("\n")}</tbody>
  </table>`;
}

/**
 * How many header rows a table opens with.
 *
 * Only those are the table's heading, and only they belong in a `<thead>` that
 * repeats when the table runs onto another page. A header row further down is
 * a row a document chose to emphasise — a totals row is the usual one — and
 * lifting it to the top would print the total above the figures it adds up.
 */
function leadingHeaders(rows: readonly TableRowNode[]): number {
  const first = rows.findIndex((row) => !row.header);

  return first === -1 ? rows.length : first;
}

/** The rows a table holds, walking through any loop that produces them. */
function tableRows(children: readonly DocumentNode[]): TableRowNode[] {
  const rows: TableRowNode[] = [];

  for (const child of children) {
    if (child.kind === "tableRow") {
      rows.push(child);
      continue;
    }

    if (child.kind === "repeat") {
      rows.push(...tableRows(child.children));
    }
  }

  return rows;
}

function renderRow(row: TableRowNode, table: TableNode): string {
  const tag = row.header ? "th" : "td";
  const cells: string[] = [];
  let column = 0;

  for (const child of row.children) {
    if (child.kind !== "tableCell") {
      continue;
    }

    const span = child.span ?? 1;
    const align = child.align ?? table.columns[column]?.align ?? "left";

    cells.push(
      `<${tag} data-node-id="${escapeHtml(child.id)}"${variantAttr(child)}${span > 1 ? ` colspan="${span}"` : ""}` +
        ` style="text-align:${align}">${child.children.map(renderNode).join("\n")}</${tag}>`,
    );
    column += span;
  }

  return `<tr data-node-id="${escapeHtml(row.id)}"${variantAttr(row)}>${cells.join("")}</tr>`;
}

/**
 * The CSS for every block style the document's theme names.
 *
 * A node says `variant="badge"`; this is where the theme's idea of a badge
 * becomes rules. A variant the theme has never heard of simply matches nothing,
 * which is why an unknown name draws as an ordinary block rather than an error.
 */
function blockCss(style: DocumentStyle): string {
  const blocks = style.blocks;

  if (!blocks) {
    return "";
  }

  return Object.entries(blocks)
    .map(([name, block]) => {
      const rules = [
        block.fill === undefined ? "" : `background: ${cssColor(block.fill)};`,
        block.color === undefined ? "" : `color: ${cssColor(block.color)};`,
        borderCss(block),
        block.paddingPt === undefined ? "" : `padding: ${block.paddingPt}pt;`,
        block.fontSizePt === undefined ? "" : `font-size: ${block.fontSizePt}pt;`,
        block.weight === undefined ? "" : `font-weight: ${block.weight === "bold" ? 700 : 400};`,
        block.transform === undefined ? "" : `text-transform: ${block.transform};`,
        block.letterSpacingEm === undefined ? "" : `letter-spacing: ${block.letterSpacingEm}em;`,
        block.bleed !== true ? "" : "margin-left: calc(-1 * var(--page-margin-left));",
        block.bleed !== true ? "" : "margin-right: calc(-1 * var(--page-margin-right));",
        block.bleed !== true ? "" : "width: auto;",
      ].filter((rule) => rule !== "");

      if (rules.length === 0) {
        return "";
      }

      const selector = cssName(name);
      const css = `      [data-variant="${selector}"] {\n` +
        rules.map((rule) => `        ${rule}`).join("\n") +
        "\n      }\n";

      // A filled block draws its own ground, so the rule that separates body
      // rows would cut across it. A band is one strip, not a row of boxes —
      // and a filled table shows its fill through cells that draw none.
      return css + (
        block.fill === undefined ? "" : `      td[data-variant="${selector}"],
      th[data-variant="${selector}"] {
        border-bottom: 0;
      }

      table[data-variant="${selector}"] td,
      table[data-variant="${selector}"] th {
        background: transparent;
        border-bottom: 0;
      }
`
      );
    })
    .join("");
}

function borderCss(block: DocumentBlockStyle): string {
  if (block.border === undefined) {
    return "";
  }

  const width = block.borderWidthPt ?? 1;
  const line = `${width}pt solid ${cssColor(block.border)}`;
  const sides = block.borderSides;

  if (!sides || sides.length === 4) {
    return `border: ${line};`;
  }

  return sides.map((side) => `border-${side}: ${line};`).join(" ");
}

/** A variant name, reduced to what is safe to put in a selector. */
function cssName(value: string): string {
  return value.replace(/[^a-zA-Z0-9_-]/g, "");
}

/** The `data-variant` attribute a node carries, when it named one. */
function variantAttr(node: DocumentNode): string {
  return node.variant === undefined ? "" : ` data-variant="${escapeHtml(cssName(node.variant))}"`;
}

function liveReloadScript(): string {
  return `<script>
    const events = new EventSource("/events");
    events.addEventListener("reload", () => window.location.reload());
  </script>`;
}

function escapeHtml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
