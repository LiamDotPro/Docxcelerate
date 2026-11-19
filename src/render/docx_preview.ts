/**
 * Settling a preview that docx-preview has just laid out.
 *
 * A preview is the packed `.docx` read back — that is the only way it cannot
 * drift from what Word shows. But docx-preview does not read everything the
 * file says: it drops a field run entirely, it looks for a table's indent
 * under an attribute that never carries it, and it wraps a picture in an
 * element a paragraph is not allowed to hold. Each one leaves the preview
 * showing something the document does not say.
 *
 * So the reading is finished here. Every function below writes back a fact the
 * packed file already declares and Word already draws — none of them invents
 * anything, and none of them is a stylesheet. The distinction matters: a
 * preview patched with CSS Word cannot reproduce is a preview that lies, and
 * the whole point of reading the file back is to stop it lying.
 *
 * Call {@linkcode settleDocxPreview} after `renderAsync` and before the
 * container is measured, screenshotted or serialised. It takes the container
 * docx-preview rendered into, so it works the same in a browser as it does in
 * a jsdom the build runs in.
 *
 * @module
 */

import type { DocumentModel, DocumentNode, ImageFloat } from "../domain/types.ts";
import { fillPageFields } from "./docx_fields.ts";
import { columnWidthMm } from "./page_size.ts";
import type { PackedParagraph, PackedTable } from "./docx_packed.ts";

export { fillPageFields } from "./docx_fields.ts";
export {
  paginateDocxPreview,
  type PaginationOptions,
  type PaginationResult,
} from "./docx_paginate.ts";
export { applyTabStops } from "./docx_tabs.ts";
export {
  previewPageStyles,
  type PreviewPageStyleOptions,
} from "./preview_styles.ts";

export {
  type ChartDrawer,
  type PreviewChart,
  type PreviewChartSeries,
  readPackedCharts,
  settleDocxPreviewCharts,
} from "./docx_preview_charts.ts";

export {
  type PackedBorderSpace,
  type PackedParagraph,
  type PackedRun,
  type PackedTable,
  type PackedTabStop,
  readPackedParagraphs,
  readPackedTables,
  readPart,
} from "./docx_packed.ts";

/** A millimetre, in points. */
const PT_PER_MM = 72 / 25.4;

/**
 * Finishes docx-preview's reading of a document it has already laid out.
 *
 * @param container The element `renderAsync` rendered into.
 * @param model The document that was packed, for the facts the DOM cannot
 * supply on its own — page size, margins, and where the body sits on the page.
 * @param packed What the file says about its own paragraphs, from
 * {@linkcode readPackedParagraphs}. Two of docx-preview's omissions can only be
 * put back from the file itself; without this they stay omitted, and the
 * preview is as complete as docx-preview left it rather than wrong.
 * @param tables What the file says about its own tables, from
 * {@linkcode readPackedTables}. One omission so far and the same rule: which
 * rows a table repeats at the top of each page it runs onto is a fact only the
 * file carries, and without it the paginator has no heading to carry.
 *
 * @example
 * ```ts
 * const bytes = new Uint8Array(await blob.arrayBuffer());
 * await renderAsync(bytes, body, head, { breakPages: true });
 * settleDocxPreview(
 *   body,
 *   model,
 *   await readPackedParagraphs(bytes),
 *   await readPackedTables(bytes),
 * );
 * ```
 */
export function settleDocxPreview(
  container: Element,
  model?: DocumentModel,
  packed?: readonly PackedParagraph[],
  tables?: readonly PackedTable[],
): void {
  if (model !== undefined) {
    pinFloatingPictures(container, model);
  }
  inlinePictureWrappers(container);
  matchStruts(container);
  addParagraphSpacing(container);
  applyDoubleBorders(container);
  liftShapeText(container);
  fillPageFields(container, model?.style?.page, model === undefined ? undefined : pageSeparators(model));
  seatFooters(container);
  hangListMarkers(container);

  if (model !== undefined) {
    applyTableIndents(container, model);
    alignPageBodies(container, model);
  }

  if (packed !== undefined) {
    applyPackedParagraphs(container, packed);
  }

  if (tables !== undefined) {
    markHeaderRows(container, tables);
  }
}

/**
 * The rows a table repeats at the top of every page, moved into a `<thead>`.
 *
 * `w:tblHeader` is a third thing docx-preview neither reads nor records: every
 * row comes out as a plain `<tr>`, so nothing in the DOM says which of them is
 * a heading and the paginator has nothing to carry onto a new sheet. Word
 * repeats them, and a preview of a two-page table with its column headings on
 * only the first page is a preview of a document nobody will receive.
 *
 * A `<thead>` rather than an attribute, because that is what the element
 * already means — "these rows head the table, and repeat" is the definition of
 * `<thead>` in HTML and of `w:tblHeader` in OOXML. Anything reading the DOM
 * afterwards, this package's paginator included, gets the fact in the form it
 * would have looked for.
 *
 * Matched by position: the nth table here is the nth `<table>` in the DOM,
 * outer tables before the ones nested inside them, which is the order both
 * lists are built in. A table has no words of its own to be matched on the way
 * a paragraph does.
 */
function markHeaderRows(container: Element, tables: readonly PackedTable[]): void {
  const drawn = [...container.querySelectorAll("table")];

  for (const [index, table] of drawn.entries()) {
    const count = tables[index]?.headerRows ?? 0;

    // A table whose every row is a heading has no body to repeat above, and
    // one already carrying a `<thead>` has been settled once already.
    if (count === 0 || count >= table.rows.length || table.tHead !== null) {
      continue;
    }

    const head = table.ownerDocument.createElement("thead");
    const rows = [...table.rows].slice(0, count);

    // Before everything else in the table, which is where a `<thead>` goes and
    // where these rows already are.
    table.insertBefore(head, table.firstChild);

    for (const row of rows) {
      head.appendChild(row);
    }
  }
}

/**
 * The DOM paragraph a packed paragraph was drawn into, found by its words.
 *
 * By text rather than by position: docx-preview renders headers and footers
 * into the same container, and the packer inserts paragraphs of its own — a
 * break carrier, a separator between two tables — so counting would drift, and
 * a drifted match writes one paragraph's appearance onto another's words. Each
 * element is claimed once, so two paragraphs that happen to say the same thing
 * take one match each.
 *
 * A paragraph that cannot be found is skipped. That is the safe direction: a
 * miss leaves the preview exactly as docx-preview drew it, where a wrong match
 * would make it say something the document never did.
 */
/**
 * A footer sits on the line the file measures it from, not above it.
 *
 * `w:pgMar/@w:footer` is the distance from the bottom of the paper to the
 * *bottom* of the footer, so Word grows a footer upward from that line.
 * docx-preview instead reserves the strip a box the height of the margin it
 * has to play with and lays the content out from the top of it — so a bar
 * shorter than its reserve floats above where Word draws it.
 *
 * Measured rather than reasoned: the invoice's footer bar is packed with a
 * distance of zero, Word's own export puts its ink one pixel off the edge of
 * the sheet, and the preview drew it six pixels up. Six pixels of white under
 * a bar that reaches every other edge of the paper is the kind of difference
 * somebody signs off on and then does not get.
 *
 * Bottom-aligning is right whichever way the sizes fall: content taller than
 * the reserve overflows upward, which is the direction Word grows it too.
 */
function seatFooters(container: Element): void {
  for (const footer of container.querySelectorAll("section.docx > footer")) {
    const style = (footer as HTMLElement).style;
    style.display = "flex";
    style.flexDirection = "column";
    style.justifyContent = "flex-end";
  }
}

/**
 * A list item's words, started where its hang says rather than wherever its
 * marker happens to end.
 *
 * Word follows a marker with a tab to the hanging indent, so the words of an
 * item's first line start exactly above the words of its second. docx-preview
 * writes the same tab — the marker is the paragraph's `::before`, and its
 * content ends in one — but into a paragraph where white space collapses, so
 * the tab draws as a space and the first line starts wherever a bullet and a
 * space come to: about 3mm left of the lines under it, for every item.
 *
 * Both halves of the repair are already in the DOM. The tab is there, and so
 * is the hang: it is the paragraph's own `text-indent`, written inline from
 * the indent the file states on it. Letting the tab be a tab and setting one
 * tab stop per hang puts the words where Word does. Read as text off the
 * attribute rather than through `style`, because a DOM without a layout
 * engine behind it does not parse every property it is handed.
 */
function hangListMarkers(container: Element): void {
  for (const item of container.querySelectorAll('p[class*="docx-num-"]')) {
    const style = (item.getAttribute("style") ?? "").replace(/;\s*$/, "");
    const hang = /text-indent:\s*-([\d.]+pt)/.exec(style)?.[1];

    if (hang !== undefined) {
      item.setAttribute("style", `${style}; white-space: pre-wrap; tab-size: ${hang};`);
    }
  }
}

function matchParagraphs(
  container: Element,
  packed: readonly PackedParagraph[],
): Array<{ packed: PackedParagraph; element: HTMLElement }> {
  const wanted = packed.filter((paragraph) => paragraph.text.trim().length > 0);
  if (wanted.length === 0) {
    return [];
  }

  const elements = [...container.querySelectorAll("p")] as HTMLElement[];
  const claimed = new Set<Element>();
  const pairs: Array<{ packed: PackedParagraph; element: HTMLElement }> = [];

  for (const paragraph of wanted) {
    const text = normalise(paragraph.text);
    const element = elements.find(
      (candidate) => !claimed.has(candidate) && normalise(candidate.textContent ?? "") === text,
    );

    if (element !== undefined) {
      claimed.add(element);
      pairs.push({ packed: paragraph, element });
    }
  }

  return pairs;
}

/** Whitespace-insensitive, because a renderer may rewrap but never rewords. */
function normalise(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

/**
 * The two things docx-preview drops on the way from the file to the DOM.
 *
 * Both are read out of the packed bytes rather than recomputed from the theme
 * that produced them. That distinction is the whole discipline of this module:
 * recomputing would be a second copy of the packer's arithmetic living beside
 * the first, and the copy that nobody opens in Word is the one free to drift.
 * What is written back here is the number in the file, which is the number
 * Word draws.
 */
function applyPackedParagraphs(container: Element, packed: readonly PackedParagraph[]): void {
  for (const pair of matchParagraphs(container, packed)) {
    applyTracking(pair.element, pair.packed);
    applyBorderPadding(pair.element, pair.packed);
  }
}

/**
 * The tracking a run declares, put back.
 *
 * docx-preview reads `w:spacing` only when its parent is `w:pPr` — the branch
 * is guarded `if (elem.localName == "pPr")` — so the identical element inside
 * `w:rPr`, which is where Word keeps character spacing, is dropped on the
 * floor. Word draws it.
 *
 * The cost is not subtle. Measured on a 7pt label tracked 0.12em, Word set it
 * 29.1mm wide and the preview 23.8mm: a fifth narrower, which is the difference
 * between a label that fits its column and one that wraps.
 *
 * Runs are matched to spans in order, since docx-preview draws one span per
 * run. Where the two counts disagree — a picture among the words, a run split
 * on a tab — every span takes the first tracking the paragraph declares, which
 * is right for the paragraphs that carry tracking at all: a label is set in one
 * face throughout, or it is not a label.
 */
function applyTracking(element: HTMLElement, packed: PackedParagraph): void {
  const tracked = packed.runs.filter((run) => run.letterSpacingPt !== undefined);
  if (tracked.length === 0) {
    return;
  }

  const spans = [...element.querySelectorAll("span")] as HTMLElement[];
  const targets = spans.length === 0 ? [element] : spans;
  const aligned = targets.length === packed.runs.length;

  targets.forEach((target, index) => {
    // Where the counts disagree, a span is matched to the run that says the
    // same words before it falls back on the first tracking: a label with an
    // untracked note after a tab is two runs and three spans, and the note
    // must not take the label's tracking.
    const same = aligned
      ? undefined
      : packed.runs.find((run) => run.text !== undefined && run.text.trim() !== "" && run.text.trim() === target.textContent?.trim());
    const spacing = aligned
      ? packed.runs[index].letterSpacingPt
      : same !== undefined
      ? same.letterSpacingPt
      : tracked[0].letterSpacingPt;

    if (spacing !== undefined) {
      target.style.letterSpacing = `${spacing}pt`;
    }
  });
}

/**
 * The gap between a paragraph's border and its words, put back.
 *
 * `w:pBdr` carries a `w:space` on each edge — the distance Word holds the text
 * off the rule — and docx-preview's `parseBorderProperties` writes only the
 * `border-*` shorthand, never the space. So a panel padded 8pt draws its border
 * tight against its own words on screen and 8pt clear of them in Word, and
 * every paragraph below it sits about 3mm too high.
 *
 * Only the edges the file actually draws a border on carry a gap, because those
 * are the only edges Word pads: `w:space` lives on the border element, so an
 * edge with no border has nowhere to record one. A block that asks for padding
 * without a border gets none here *and* none in Word — a limitation of the
 * format rather than of the reading, and the reason a filled panel that must
 * hold its words off its own edge has to be a table cell.
 *
 * The two axes are not the same shape, which is the part only a measurement
 * tells you. Measured on an 8pt-padded panel: Word moves the text *down* from a
 * top border, so the vertical gap is padding as CSS means it; but it does not
 * move the text *in* from a left border — it moves the border out into the
 * margin and leaves the words on the column. So the horizontal gap is padding
 * with a matching negative margin, which draws the same picture: the rule steps
 * outward, the text does not move.
 */
function applyBorderPadding(element: HTMLElement, packed: PackedParagraph): void {
  const sides = [
    ["top", "Top"],
    ["right", "Right"],
    ["bottom", "Bottom"],
    ["left", "Left"],
  ] as const;

  for (const [side, capitalised] of sides) {
    const gap = packed.borderSpacePt[side];
    if (gap === undefined || gap === 0) {
      continue;
    }

    element.style[`padding${capitalised}`] = `${gap}pt`;

    // Left and right: the border steps out, the text stays where it was. The
    // paragraph's own indent is already on the inline margin, and a physical
    // `margin-left` would replace it rather than add to it.
    if (side === "left" || side === "right") {
      const indent = element.style[side === "left" ? "marginInlineStart" : "marginInlineEnd"] || "0pt";
      element.style[`margin${capitalised}`] = `calc(${indent} - ${gap}pt)`;
    }
  }
}

/**
 * The box docx-preview puts a picture in, made something a paragraph can hold.
 *
 * A drawing is wrapped in a `<div>`, and docx-preview appends it to the `<p>`
 * its run belongs to. In a DOM built by hand that stands; serialised to HTML
 * and parsed again it does not, because a `<div>` may not sit inside a `<p>`.
 * The parser closes the paragraph at the picture, strands the words that
 * followed it, and leaves an empty paragraph where the close tag lands — one
 * taking the document's default leading rather than the line the file asked
 * for. Measured on an invoice, that turned a one-line footer bar 29px deep in
 * Word into 73px in the preview.
 *
 * The wrapper does the job of a `<span>` set `inline-block`, and a span is what
 * a paragraph may contain. Nothing about the layout changes; what changes is
 * that the layout survives being written down.
 */
export function inlinePictureWrappers(container: Element): void {
  for (const wrapper of [...container.querySelectorAll("p div")]) {
    const span = wrapper.ownerDocument.createElement("span");

    span.setAttribute("style", wrapper.getAttribute("style") ?? "");
    while (wrapper.firstChild !== null) {
      span.appendChild(wrapper.firstChild);
    }
    wrapper.replaceWith(span);
  }
}


/** What each of the document's page numbers writes between its two fields. */
function pageSeparators(model: DocumentModel): string[] {
  const found = new Set<string>();
  const walk = (nodes: readonly DocumentNode[] | undefined) => {
    for (const node of nodes ?? []) {
      if (node.kind === "pageNumber" && (node.format ?? "currentOfTotal") === "currentOfTotal") {
        found.add(node.separator ?? " / ");
      }
      if ("children" in node) walk(node.children);
    }
  };
  for (const part of [model.header, model.footer, model.firstHeader, model.firstFooter, model.evenHeader, model.evenFooter]) {
    walk(part);
  }
  return [...found];
}

/**
 * Every floating picture, put where the file pins it: on the page, not in the line.
 *
 * docx-preview reads a `wp:anchor` that does not wrap as a zero-sized box
 * nudged by the anchor's offsets from wherever its paragraph happens to be —
 * the offsets are from the paper's corner in the file, so the picture lands a
 * paragraph's distance too low — and it reads `behindDoc` and then drops it,
 * so a cover photograph paints over the title set on it. Word draws both from
 * the page.
 *
 * The pictures are matched to the model's floating images in order, and the
 * position written back is the model's own, so nothing here is measured or
 * guessed. Behind means behind the words and in front of the sheet, which is
 * what an isolated stacking context on the sheet gives a negative z-index.
 */
function pinFloatingPictures(container: Element, model: DocumentModel): void {
  const floats: ImageFloat[] = [];
  const walk = (nodes: readonly DocumentNode[]) => {
    for (const node of nodes) {
      if (node.kind === "image" && node.float !== undefined) floats.push(node.float);
      // A picture set in a paragraph is carried beside its text, not as a child.
      if (node.kind === "paragraph") {
        for (const inline of node.inlineImages ?? []) {
          if (inline.image.float !== undefined) floats.push(inline.image.float);
        }
      }
      if ("children" in node) walk(node.children);
    }
  };
  walk(model.nodes);
  if (floats.length === 0) return;

  const anchors = [...container.querySelectorAll<HTMLElement>("section.docx article div")]
    .filter((box) => box.style.width === "0px" && box.style.height === "0px" && box.querySelector("img"));

  anchors.slice(0, floats.length).forEach((box, index) => {
    const float = floats[index]!;
    const sheet = box.closest<HTMLElement>("section.docx");
    if (sheet === null) return;

    sheet.style.position = "relative";
    sheet.style.isolation = "isolate";
    box.style.position = "absolute";
    box.style.left = `${float.xMm}mm`;
    box.style.top = `${float.yMm}mm`;
    box.style.zIndex = float.behind === true ? "-1" : "1";
    sheet.prepend(box);
  });
}

/**
 * The space between two paragraphs, added up the way Word adds it.
 *
 * Word leaves a paragraph's space after *and* the next one's space before.
 * docx-preview writes both as CSS margins, and CSS collapses two touching
 * margins into the larger of them — so a label 7pt above a point padded 8pt
 * below its own top came out 8pt apart on screen and 15pt apart in Word, and
 * every such pair let the preview run shorter than the file. Folding the one
 * margin into the other gives CSS's collapse the sum to keep.
 */
function addParagraphSpacing(container: Element): void {
  for (const paragraph of container.querySelectorAll<HTMLElement>("p")) {
    const before = paragraph.previousElementSibling as HTMLElement | null;
    if (before === null || before.tagName !== "P") continue;
    const above = Number.parseFloat(before.style.marginBottom) || 0;
    const own = Number.parseFloat(paragraph.style.marginTop) || 0;
    if (above > 0 && own > 0) paragraph.style.marginTop = `${above + own}pt`;
  }
}

/**
 * Each paragraph set at the size of its own words.
 *
 * docx-preview writes a run's size on the run's span and leaves the paragraph
 * at the page's default size. CSS then lays every line out around a strut in
 * that default face, and when the run is a different size, aligning the two on
 * one baseline makes the line box a pixel or so deeper than the leading the
 * file states — on a page of labelled bars, a pixel a row. Word has no strut;
 * a line is exactly the height its spacing says. Where every run of a
 * paragraph is one size, that size is the paragraph's, which is what puts the
 * preview's lines back at the file's height.
 */
function matchStruts(container: Element): void {
  for (const paragraph of container.querySelectorAll<HTMLElement>("p")) {
    if (paragraph.style.fontSize !== "") continue;
    const sizes = new Set(
      [...paragraph.querySelectorAll<HTMLElement>(":scope > span")]
        .map((span) => span.style.fontSize)
        .filter((size) => size !== ""),
    );
    if (sizes.size === 1) paragraph.style.fontSize = [...sizes][0]!;
  }
}

/**
 * Widens every double border to the three widths Word draws it across.
 *
 * Word draws a double border as two lines of the border's width with a gap of
 * the same width between them, so three widths overall, which is exactly the
 * geometry of CSS `double`. docx-preview writes the file's single width, which
 * CSS needs at least 3px to show as two lines, so a 1pt rule would read as one
 * solid line. Widening it also keeps the rows below in step with Word, where
 * the full width takes up layout space too.
 */
function applyDoubleBorders(container: Element): void {
  for (const el of container.querySelectorAll<HTMLElement>("p, td")) {
    for (const side of ["Top", "Right", "Bottom", "Left"] as const) {
      if (el.style[`border${side}Style`] === "double") {
        // docx-preview writes border widths in pt, from LengthUsage.Border.
        el.style[`border${side}Width`] = `${parseFloat(el.style[`border${side}Width`]) * 3}pt`;
      }
    }
  }
}

/**
 * A shape's words, moved somewhere SVG will paint them.
 *
 * docx-preview renders a VML shape by walking it and nesting each child inside
 * the element it drew for the parent — so a `v:rect` holding a `v:textbox`
 * comes out as an SVG `<rect>` holding a `<foreignObject>`. That is faithful
 * to the file's nesting and invalid as SVG: `<rect>` is a shape, its content
 * model holds no rendered children, and a browser paints none of them.
 * Measured in Chrome, the words inside one report a bounding box of 0 by 0
 * while the same `<foreignObject>` moved out beside the rect reports 92 by 17.
 *
 * So a shape drew as a filled block with nothing on it, which is the worst
 * shape of all for a preview bug: it looks deliberate. Word draws the text.
 *
 * Moving the node is the whole fix. Both elements come from the packed file,
 * both are drawn by Word, and their order is already right — a `<foreignObject>`
 * after the `<rect>` paints on top of it, which is where the file says the
 * words are. Nothing here invents a size, a colour or a position.
 */
export function liftShapeText(container: Element): void {
  // Every shape element SVG will not paint children of. `rect` and `ellipse`
  // are what docx-preview draws a `v:rect` and a `v:oval` as; a `<g>`, which
  // is what it draws a `v:shape` as, is a container and needs no help.
  for (const shape of [...container.querySelectorAll("svg rect, svg ellipse")]) {
    for (const text of [...shape.querySelectorAll("foreignObject")]) {
      shape.insertAdjacentElement("afterend", text);
    }
  }
}

/**
 * The body set where `w:vAlign` puts it on the page.
 *
 * docx-preview reads the section's size, margins and furniture and skips its
 * vertical alignment, so a centred certificate drew at the top of its sheet.
 * Its own stylesheet already lays a page out as a flex column whose article
 * has `margin-bottom: auto` (that is what holds the footer down), so taking
 * up the slack above the article as well is all centring is.
 */
function alignPageBodies(container: Element, model: DocumentModel): void {
  const valign = model.style?.page?.valign;
  if (valign !== "center" && valign !== "bottom") return;

  for (const article of container.querySelectorAll("section.docx > article")) {
    const style = (article as HTMLElement).style;
    style.marginTop = "auto";
    if (valign === "bottom") style.marginBottom = "0px";
  }
}

/**
 * The indent a bleeding table declares, put back.
 *
 * docx-preview reads a table's indent with `parseIndentation`, which looks for
 * a `w:left` attribute — and `w:tblInd` does not carry one, its value is in
 * `w:w`. So the indent is parsed and then dropped, and a table the file says
 * reaches the paper's edge stops at the margin instead. Word honours it.
 *
 * A bleeding table is known by its own width: docx-preview writes the table's
 * declared width onto the element, and nothing but a bleed makes a table wider
 * than the text column it stands in. That keeps this from needing to know
 * which variant a table was given.
 */
export function applyTableIndents(container: Element, model: DocumentModel): void {
  const margins = model.style?.page?.margins;
  if (margins === undefined) {
    return;
  }

  const columnPt = columnWidthMm(model.style!.page) * PT_PER_MM;

  for (const table of [...container.querySelectorAll("table")]) {
    const element = table as HTMLElement;
    const declared = Number.parseFloat(element.style.width);

    // A point of slack: the declared width is written back from twips, so a
    // table sitting exactly on the column can round either way.
    if (!Number.isFinite(declared) || declared <= columnPt + 1) {
      continue;
    }

    element.style.marginInlineStart = `-${margins.leftMm}mm`;
    element.style.marginInlineEnd = `-${margins.rightMm}mm`;
  }
}
