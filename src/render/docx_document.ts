/**
 * Packing a finished document model into a Word file.
 *
 * @module
 */

import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HeadingLevel,
  ImageRun,
  Packer,
  PageBreak,
  PageNumber,
  PageOrientation,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import type {
  DocumentBlockStyle,
  DocumentModel,
  DocumentNode,
  DocumentStyle,
  DocumentTextBlockStyle,
  ImageNode,
  PageNumberNode,
  TableAlign,
  TableCellNode,
  TableColumn,
  TableNode,
  TableRowNode,
} from "../domain/types.ts";
import { cleanMinimalDocumentStyle } from "../project/style.ts";
import { imageSourceOf, isSvg, rasterTypeOf } from "./image_source.ts";

/**
 * Lays a document model out as a `docx` document — page size, margins, styles
 * and all the nodes.
 *
 * @param doc The finished document. Its style falls back to the clean minimal
 * preset when it has none.
 * @returns The `docx` document, ready to pack.
 */
export function createDocxDocument(doc: DocumentModel): Document {
  const style = doc.style ?? cleanMinimalDocumentStyle;

  return new Document({
    styles: createDocxStyles(style),
    sections: [
      {
        properties: {
          page: {
            size: {
              width: pageWidthTwips(style),
              height: pageHeightTwips(style),
              orientation: style.page.orientation === "landscape"
                ? PageOrientation.LANDSCAPE
                : PageOrientation.PORTRAIT,
            },
            margin: {
              top: mmToTwips(style.page.margins.topMm),
              right: mmToTwips(style.page.margins.rightMm),
              bottom: mmToTwips(style.page.margins.bottomMm),
              left: mmToTwips(style.page.margins.leftMm),
            },
          },
        },
        headers: doc.header
          ? { default: new Header({ children: paragraphsOf(doc.header, style) }) }
          : undefined,
        footers: doc.footer
          ? { default: new Footer({ children: paragraphsOf(doc.footer, style) }) }
          : undefined,
        children: [
          ...(style.showTitle === false ? [] : [
            new Paragraph({ text: doc.title, heading: HeadingLevel.TITLE }),
          ]),
          ...doc.nodes.flatMap((node) => renderNode(node, style)),
        ],
      },
    ],
  });
}

/**
 * Word's headers and footers take paragraphs and tables, and nothing else.
 *
 * A stray node of another kind is dropped rather than packed as text, because
 * running furniture is the one place a surprise is printed on every page.
 */
function paragraphsOf(nodes: readonly DocumentNode[], style: DocumentStyle): (Paragraph | Table)[] {
  const content = nodes.flatMap((node) => renderNode(node, style, true));

  return content.length === 0 ? [new Paragraph({})] : content;
}

/**
 * Packs a document model into a `.docx` file.
 *
 * @param doc The finished document.
 * @returns The file, as a blob.
 *
 * @example
 * ```ts
 * const blob = await createDocxBlob(model);
 * await Deno.writeFile("document.docx", new Uint8Array(await blob.arrayBuffer()));
 * ```
 */
export async function createDocxBlob(doc: DocumentModel): Promise<Blob> {
  return await Packer.toBlob(createDocxDocument(doc));
}

/**
 * Packs a document model into `.docx` bytes.
 *
 * The counterpart to {@linkcode createDocxBlob}, for writing straight to a file
 * or a response body rather than handing a blob to a browser.
 *
 * @param doc The finished document.
 * @returns The file's bytes.
 */
export async function renderDocxBytes(doc: DocumentModel): Promise<Uint8Array> {
  return await Packer.toBuffer(createDocxDocument(doc));
}

function renderNode(
  node: DocumentNode,
  style: DocumentStyle,
  furniture = false,
): (Paragraph | Table)[] {
  if (node.kind === "section") {
    return [
      new Paragraph({
        text: node.title ?? node.id,
        heading: HeadingLevel.HEADING_1,
      }),
      ...node.children.flatMap((child) => renderNode(child, style, furniture)),
    ];
  }

  if (node.kind === "paragraph") {
    const block = blockOf(style, node.variant);

    return [
      new Paragraph({
        children: [new TextRun({ text: node.text ?? "", ...blockRun(block, style) })],
        shading: block?.fill === undefined
          ? undefined
          : { type: ShadingType.CLEAR, fill: block.fill },
        border: blockBorders(block),
        indent: blockIndent(block, style),
        spacing: {
          after: ptToTwips(style.paragraph.spacingAfterPt),
          line: Math.round(style.typography.bodyLineHeight * 240),
        },
      }),
    ];
  }

  if (node.kind === "image") {
    return [new Paragraph({ children: [imageRunOf(node)] })];
  }

  if (node.kind === "graph") {
    return [
      new Paragraph({
        children: [
          new TextRun(`[${node.graphType} graph: ${node.caption ?? node.placeholder ?? node.id}]`),
        ],
      }),
    ];
  }

  if (node.kind === "table") {
    return [renderTable(node, style, furniture)];
  }

  if (node.kind === "pageBreak") {
    return [new Paragraph({ children: [new PageBreak()] })];
  }

  // Word counts the pages itself, as fields it re-evaluates whenever the
  // document is repaginated. A number written in at build time would be a
  // number that stops being true the moment an engine writes a longer
  // paragraph into a dynamic node.
  if (node.kind === "pageNumber") {
    return [new Paragraph({ children: pageNumberRuns(node) })];
  }

  // A built document has already walked its loops, so this is a published one
  // being packed directly. The body is all there is to show: one pass, standing
  // for however many the request will ask for.
  //
  // A row or a cell reaching here is one outside the table it belongs to, and
  // its content is still content — printing it beats dropping it silently.
  if (node.kind === "repeat" || node.kind === "tableRow" || node.kind === "tableCell") {
    return node.children.flatMap((child) => renderNode(child, style, furniture));
  }

  return [
    new Paragraph({
      children: [new TextRun(node.title ?? "Table of contents")],
    }),
  ];
}

/**
 * The block style a variant names, when the theme has one.
 *
 * A name the theme has never heard of draws as an ordinary block. That is
 * deliberate: swapping in a theme that knows fewer variants should restyle a
 * document, not refuse to pack it.
 */
function blockOf(
  style: DocumentStyle,
  variant: string | undefined,
): DocumentBlockStyle | undefined {
  return variant === undefined ? undefined : style.blocks?.[variant];
}

/**
 * How a block's text is set, as Word measures it.
 *
 * The same fields the screen reads, so a variant means one thing rather than
 * two. Character spacing is in twentieths of a point rather than ems, so the
 * font size is what turns one into the other.
 */
function blockRun(
  block: DocumentBlockStyle | undefined,
  style: DocumentStyle,
): {
  bold?: boolean;
  size?: number;
  color?: string;
  allCaps?: boolean;
  characterSpacing?: number;
} {
  if (!block) {
    return {};
  }

  const sizePt = block.fontSizePt ?? style.typography.bodySizePt;

  return {
    bold: block.weight === undefined ? undefined : block.weight === "bold",
    size: block.fontSizePt === undefined ? undefined : ptToHalfPoints(block.fontSizePt),
    color: block.color,
    allCaps: block.transform === "uppercase" || undefined,
    characterSpacing: block.letterSpacingEm === undefined
      ? undefined
      : Math.round(block.letterSpacingEm * sizePt * 20),
  };
}

/**
 * The borders a block draws, as Word measures them.
 *
 * Width is in eighths of a point and the gap to the text in whole points, both
 * of which Word caps — a border it cannot draw is worse than one drawn a shade
 * thinner, so the numbers are clamped rather than passed through.
 */
function blockBorders(block: DocumentBlockStyle | undefined) {
  if (!block?.border) {
    return undefined;
  }

  const edge = {
    style: BorderStyle.SINGLE,
    color: block.border,
    size: Math.max(1, Math.round((block.borderWidthPt ?? 1) * 8)),
    space: Math.min(31, Math.round(block.paddingPt ?? 0)),
  };
  const sides = block.borderSides ?? ["top", "right", "bottom", "left"];

  return Object.fromEntries(sides.map((side) => [side, edge]));
}

/**
 * The hairline a row is separated from the next one by.
 *
 * One line under each body cell, in the palette's rule colour, and nothing
 * else: a heading already has its fill to set it apart, a cell that draws its
 * own ground would be cut across by a rule at its foot, and running furniture
 * is not prose being separated into rows.
 */
function separatorBorder(
  row: TableRowNode,
  fill: string | undefined,
  furniture: boolean,
  style: DocumentStyle,
) {
  if (row.header === true || fill !== undefined || furniture) {
    return undefined;
  }

  return {
    bottom: {
      style: BorderStyle.SINGLE,
      color: style.palette?.rule ?? "9AA6B8",
      // A screen pixel is three quarters of a point, and Word counts a border
      // in eighths of one.
      size: 6,
      space: 0,
    },
  };
}

/**
 * What a bleeding block indents by, so it reaches past the margins.
 *
 * Word measures a paragraph from the margin, so reaching outside one is a
 * negative indent — the same distance the margin is wide, which puts the block
 * flush with the edge of the sheet.
 */
function blockIndent(block: DocumentBlockStyle | undefined, style: DocumentStyle) {
  if (block?.bleed !== true) {
    return undefined;
  }

  return {
    left: -mmToTwips(style.page.margins.leftMm),
    right: -mmToTwips(style.page.margins.rightMm),
  };
}

/**
 * The picture a Word file embeds, or the note that stands in for one.
 *
 * Packing means embedding bytes, and the only bytes a model carries are the
 * ones in a `data:` URI. A path or an http URL is not reached for: this runs
 * wherever the document is written, which is not where the file was, and a
 * renderer that quietly fetched would produce a document that packs on one
 * machine and not another.
 */
function imageRunOf(node: ImageNode): ImageRun | TextRun {
  const source = imageSourceOf(node.path);
  const described = node.alt ?? node.placeholder ?? node.id;

  if (source.kind !== "data") {
    return new TextRun(`[image: ${described}]`);
  }

  const transformation = {
    width: ptToPx(node.width ?? 120),
    height: ptToPx(node.height ?? node.width ?? 120),
  };
  const altText = { name: node.id, description: described, title: described };
  const raster = rasterTypeOf(source.mediaType);

  if (raster) {
    return new ImageRun({ type: raster, data: source.bytes, transformation, altText });
  }

  if (isSvg(source.mediaType)) {
    const fallback = imageSourceOf(node.fallbackPath);
    const fallbackType = fallback.kind === "data" ? rasterTypeOf(fallback.mediaType) : undefined;

    // Word will not take an SVG on its own. Without a raster to fall back to
    // there is nothing honest to embed, so the note says what it was rather
    // than the document showing an empty frame.
    if (fallback.kind === "data" && fallbackType) {
      return new ImageRun({
        type: "svg",
        data: source.bytes,
        fallback: { type: fallbackType, data: fallback.bytes },
        transformation,
        altText,
      });
    }
  }

  return new TextRun(`[image: ${described}]`);
}

/** Points to pixels, the unit Word measures a picture in. */
function ptToPx(points: number): number {
  return Math.round(points * (96 / 72));
}

/** The runs a page number becomes — Word's own fields, not a printed digit. */
function pageNumberRuns(node: PageNumberNode): TextRun[] {
  const format = node.format ?? "currentOfTotal";

  if (format === "current") {
    return [new TextRun({ children: [PageNumber.CURRENT] })];
  }

  if (format === "total") {
    return [new TextRun({ children: [PageNumber.TOTAL_PAGES] })];
  }

  return [
    new TextRun({ children: [PageNumber.CURRENT] }),
    new TextRun(node.separator ?? " / "),
    new TextRun({ children: [PageNumber.TOTAL_PAGES] }),
  ];
}

/**
 * A real Word table, so the columns are columns rather than tabbed text.
 *
 * The widths are worked out here rather than left to Word, because a column
 * declared in millimetres has to come out that wide on the page — an invoice's
 * money column lining up is the whole reason it was given a width. Whatever
 * the fixed columns leave is shared between the `"auto"` ones.
 */
function renderTable(node: TableNode, style: DocumentStyle, furniture: boolean): Table {
  const widths = columnWidths(node.columns, style);
  const rows = tableRows(node.children);

  // Only the rows a table opens with are its heading. `tableHeader` is what
  // makes Word repeat a row at the top of every page the table runs onto, so
  // setting it on a totals row further down would print the total on each page
  // above the figures it adds up.
  const lead = rows.findIndex((row) => !row.header);
  const headers = lead === -1 ? rows.length : lead;

  return new Table({
    columnWidths: widths,
    width: { size: widths.reduce((total, width) => total + width, 0), type: WidthType.DXA },
    // Word's own default is a black grid around every cell, which no theme
    // ever asked for and the screen never draws. The lines a table shows are
    // the ones decided below, one per row and in the palette's rule colour.
    borders: gridlessBorders,
    rows: rows.map((row, index) => renderRow(row, node, style, index < headers, furniture)),
  });
}

/** No lines at all, so the only ones drawn are the ones a cell asks for. */
const gridlessBorders = {
  top: { style: BorderStyle.NONE, size: 0, color: "auto" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "auto" },
  left: { style: BorderStyle.NONE, size: 0, color: "auto" },
  right: { style: BorderStyle.NONE, size: 0, color: "auto" },
  insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "auto" },
  insideVertical: { style: BorderStyle.NONE, size: 0, color: "auto" },
};

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

function renderRow(
  row: TableRowNode,
  table: TableNode,
  style: DocumentStyle,
  repeats: boolean,
  furniture: boolean,
): TableRow {
  const cells: TableCell[] = [];
  let column = 0;

  for (const child of row.children) {
    if (child.kind !== "tableCell") {
      continue;
    }

    const span = child.span ?? 1;
    const align = child.align ?? table.columns[column]?.align ?? "left";

    // A cell's own variant wins over its row's, which wins over the table's —
    // the narrower statement is the more deliberate one.
    const block = blockOf(style, child.variant) ?? blockOf(style, row.variant) ??
      blockOf(style, table.variant);
    // The accent bar over a heading is what a header row looks like when the
    // theme has not named one. Once it has — for the cell, its row or its
    // table — that naming is the whole answer, and the default steps aside
    // rather than showing through the half of it the theme left unsaid.
    const fill = block?.fill ??
      (row.header && block === undefined ? style.palette?.accent ?? "1F2933" : undefined);
    // The theme's padding, or the room a cell is given when it says nothing —
    // the same room the screen leaves, since a column that lines up in the
    // preview and not in Word is the preview lying about the document.
    const padding = block?.paddingPt;
    const inset = padding === undefined
      ? { vertical: row.header ? 6 : 5, horizontal: 8 }
      : { vertical: padding, horizontal: padding };

    cells.push(
      new TableCell({
        columnSpan: span > 1 ? span : undefined,
        shading: fill === undefined ? undefined : { type: ShadingType.CLEAR, fill },
        borders: blockBorders(block) ?? separatorBorder(row, fill, furniture, style),
        margins: {
          marginUnitType: WidthType.DXA,
          top: ptToTwips(inset.vertical),
          bottom: ptToTwips(inset.vertical),
          left: ptToTwips(inset.horizontal),
          right: ptToTwips(inset.horizontal),
        },
        children: cellContent(child, align, row.header === true, block, style, furniture),
      }),
    );
    column += span;
  }

  return new TableRow({ tableHeader: repeats, children: cells });
}

/**
 * What one cell prints.
 *
 * A cell's paragraphs are rebuilt rather than taken from the ordinary node
 * renderer, which knows nothing about the column the cell landed in: the
 * alignment is the column's, and the spacing that belongs between blocks of
 * prose does not belong inside a cell.
 *
 * Word will not accept an empty cell, so a cell with nothing in it still gets
 * a paragraph — an invoice line with no note is a blank box, not a missing one.
 */
function cellContent(
  cell: TableCellNode,
  align: TableAlign,
  header: boolean,
  block: DocumentBlockStyle | undefined,
  style: DocumentStyle,
  furniture: boolean,
): (Paragraph | Table)[] {
  const content = cell.children.flatMap((child): (Paragraph | Table)[] => {
    // A picture in a cell is a picture, not a paragraph of prose: Word gives
    // it a line and a space after out of the body style, which drops a
    // letterhead's mark below the name beside it and makes a one-line footer
    // bar two lines deep.
    if (child.kind === "image") {
      return [
        new Paragraph({
          alignment: alignments[align],
          spacing: { after: 0, line: 240 },
          children: [imageRunOf(child)],
        }),
      ];
    }

    if (child.kind !== "paragraph") {
      return renderNode(child, style, furniture);
    }

    // The cell's own block, then the paragraph's — a muted note inside a
    // filled cell is set by its own variant, not the cell's.
    const inner = blockOf(style, child.variant);
    const run = { ...blockRun(block, style), ...blockRun(inner, style) };
    const heading: Partial<ReturnType<typeof headingRun>> = header
      ? headingRun(block, style)
      : {};
    const size = run.size ?? heading.size;

    return [
      new Paragraph({
        alignment: alignments[align],
        spacing: { after: 0, line: Math.round(style.typography.bodyLineHeight * 240) },
        children: [
          new TextRun({
            text: child.text ?? "",
            bold: run.bold ?? heading.bold,
            color: run.color ?? heading.color,
            size,
            allCaps: run.allCaps ?? heading.allCaps,
            characterSpacing: run.characterSpacing ??
              (heading.trackedEm === undefined
                ? undefined
                : trackingOf(heading.trackedEm, size, style)),
          }),
        ],
      }),
    ];
  });

  return content.length === 0 ? [new Paragraph({})] : content;
}

/**
 * How a heading row is set when the theme has not said otherwise.
 *
 * Small tracked capitals, reversed out of the accent — the same face the
 * screen gives a `<th>`, written here so that the two agree. Anything the
 * theme does name for the row wins over all of it, which is why the ink and
 * the size are only offered when no variant is in play.
 */
function headingRun(
  block: DocumentBlockStyle | undefined,
  style: DocumentStyle,
): { bold: boolean; allCaps: boolean; trackedEm: number; color?: string; size?: number } {
  return {
    bold: true,
    allCaps: true,
    trackedEm: 0.1,
    color: block === undefined ? style.palette?.page ?? "FFFFFF" : undefined,
    size: block === undefined
      ? ptToHalfPoints(style.typography.bodySizePt * 0.72)
      : undefined,
  };
}

/** Tracking in ems, as the twentieths of a point Word counts it in. */
function trackingOf(em: number, size: number | undefined, style: DocumentStyle): number {
  const sizePt = size === undefined ? style.typography.bodySizePt : size / 2;

  return Math.round(em * sizePt * 20);
}

const alignments: Record<TableAlign, (typeof AlignmentType)[keyof typeof AlignmentType]> = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
};

/**
 * Turns declared column widths into the twips Word lays a table out in.
 *
 * A width in millimetres is honoured exactly. `"auto"` shares out what is left
 * of the text column, which is the width the page has once its margins are off.
 */
function columnWidths(columns: readonly TableColumn[], style: DocumentStyle): number[] {
  const available = pageWidthTwips(style) -
    mmToTwips(style.page.margins.leftMm + style.page.margins.rightMm);
  const fixed = columns.map((column) =>
    typeof column.width === "number" ? mmToTwips(column.width) : 0
  );
  const autos = fixed.filter((width) => width === 0).length;
  const spare = Math.max(available - fixed.reduce((total, width) => total + width, 0), 0);
  const each = autos > 0 ? Math.floor(spare / autos) : 0;

  return fixed.map((width) => (width === 0 ? each : width));
}

function createDocxStyles(style: DocumentStyle) {
  return {
    default: {
      document: {
        run: {
          font: style.typography.bodyFont,
          size: ptToHalfPoints(style.typography.bodySizePt),
          color: style.typography.color,
        },
        paragraph: {
          spacing: {
            after: ptToTwips(style.paragraph.spacingAfterPt),
            line: Math.round(style.typography.bodyLineHeight * 240),
          },
        },
      },
      title: {
        run: {
          font: style.typography.headingFont,
          size: ptToHalfPoints(style.title.fontSizePt),
          bold: style.title.weight === "bold",
          color: headingColor(style, style.title),
          allCaps: style.title.transform === "uppercase",
        },
        paragraph: {
          spacing: {
            before: ptToTwips(style.title.spacingBeforePt),
            after: ptToTwips(style.title.spacingAfterPt),
          },
        },
      },
      heading1: {
        run: {
          font: style.typography.headingFont,
          size: ptToHalfPoints(style.sectionHeading.fontSizePt),
          bold: style.sectionHeading.weight === "bold",
          color: headingColor(style, style.sectionHeading),
          allCaps: style.sectionHeading.transform === "uppercase",
        },
        paragraph: {
          spacing: {
            before: ptToTwips(style.sectionHeading.spacingBeforePt),
            after: ptToTwips(style.sectionHeading.spacingAfterPt),
          },
        },
      },
    },
  };
}

/**
 * The ink a heading is set in: its own colour, then the palette's, then the
 * body colour.
 *
 * The last fallback is what keeps a document written before themes existed
 * looking exactly as it did — no palette means headings stay body-coloured,
 * which is what the renderer did when there was nothing else to ask.
 */
function headingColor(style: DocumentStyle, block: DocumentTextBlockStyle): string {
  return block.color ?? style.palette?.heading ?? style.typography.color;
}

function pageWidthTwips(style: DocumentStyle): number {
  return mmToTwips(style.page.size === "A4" ? 210 : 215.9);
}

function pageHeightTwips(style: DocumentStyle): number {
  return mmToTwips(style.page.size === "A4" ? 297 : 279.4);
}

function mmToTwips(value: number): number {
  return Math.round(value * 56.6929133858);
}

function ptToTwips(value: number): number {
  return Math.round(value * 20);
}

function ptToHalfPoints(value: number): number {
  return Math.round(value * 2);
}
