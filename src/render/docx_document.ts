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
  LeaderType,
  LineRuleType,
  Packer,
  PageBreak,
  PageNumber,
  PageOrientation,
  Paragraph,
  ShadingType,
  Tab,
  ImportedXmlComponent,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TabStopType,
  TextRun,
  VerticalAlign,
  WidthType,
} from "docx";
import type { IParagraphStyleOptions, IParagraphStylePropertiesOptions } from "docx";
import type {
  DocumentBlockStyle,
  DocumentModel,
  DocumentNode,
  DocumentStyle,
  DocumentTextBlockStyle,
  GraphNode,
  ImageNode,
  PageNumberNode,
  ParagraphNode,
  ShapeNode,
  TableAlign,
  TableCellNode,
  TableColumn,
  TableNode,
  TableRowNode,
  TextAlign,
} from "../domain/types.ts";
import { cleanMinimalDocumentStyle } from "../project/style.ts";
import { hasChartData } from "./chart_part.ts";
import { addChartsToPackage, chartRelationshipToken } from "./docx_charts.ts";
import { imageSourceOf, isSvg, rasterTypeOf } from "./image_source.ts";

/**
 * How far a running strip stands from the paper when a document says nothing.
 *
 * 12.5mm, which is what Word's own default template uses and what documents
 * packed before `headerMm` existed already got. Changing it would move the
 * furniture of every document that never asked for it to move.
 */
const DEFAULT_FURNITURE_MM = 12.5;

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
  // First-page furniture is Word's title page: `w:titlePg` plus a `first`
  // header or footer part. An empty array still counts — it is the document
  // saying the first page shows nothing where the others show the strip.
  const titlePage = doc.firstHeader !== undefined || doc.firstFooter !== undefined;
  // Naming either even strip turns the whole document over to two kinds of
  // page. It is a document setting rather than a section one, which is why it
  // sits out here beside the styles rather than in the section below.
  const evenAndOdd = doc.evenHeader !== undefined || doc.evenFooter !== undefined;

  return new Document({
    styles: createDocxStyles(style),
    evenAndOddHeaderAndFooters: evenAndOdd || undefined,
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
              // Where the running strips stand, measured from the paper rather
              // than from the margin. Stated rather than left to the packing
              // library's own default, which was 708 twips for every document
              // whatever its margins were — a number no document had chosen and
              // none could change.
              header: mmToTwips(style.page.headerMm ?? DEFAULT_FURNITURE_MM),
              footer: mmToTwips(style.page.footerMm ?? DEFAULT_FURNITURE_MM),
            },
          },
          titlePage: titlePage || undefined,
        },
        // `w:titlePg` is one switch for both strips: once it is on, Word takes
        // page one's header AND footer from the `first` parts and shows
        // nothing where a part is missing. So a document that names only a
        // first header would silently lose its footer off page one — absent
        // means "the first page is like every other", which is the default
        // part repeated, not an empty one.
        headers: doc.header || doc.firstHeader || doc.evenHeader
          ? {
            default: doc.header
              ? new Header({ children: paragraphsOf(doc.header, style) })
              : undefined,
            first: doc.firstHeader
              ? new Header({ children: paragraphsOf(doc.firstHeader, style) })
              : titlePage && doc.firstHeader === undefined && doc.header
              ? new Header({ children: paragraphsOf(doc.header, style) })
              : undefined,
            // The verso's own strip. Once two kinds of page are on, `header` is
            // what a right-hand page shows and this is what a left-hand one
            // does — so a document that named only an even header still wants
            // the running one on its rectos, which is the `default` above.
            even: doc.evenHeader
              ? new Header({ children: paragraphsOf(doc.evenHeader, style) })
              : undefined,
          }
          : undefined,
        footers: doc.footer || doc.firstFooter || doc.evenFooter
          ? {
            default: doc.footer
              ? new Footer({ children: paragraphsOf(doc.footer, style) })
              : undefined,
            first: doc.firstFooter
              ? new Footer({ children: paragraphsOf(doc.firstFooter, style) })
              : titlePage && doc.firstFooter === undefined && doc.footer
              ? new Footer({ children: paragraphsOf(doc.footer, style) })
              : undefined,
            even: doc.evenFooter
              ? new Footer({ children: paragraphsOf(doc.evenFooter, style) })
              : undefined,
          }
          : undefined,
        children: [
          ...(style.showTitle === false ? [] : [
            new Paragraph({ text: doc.title, heading: HeadingLevel.TITLE }),
          ]),
          ...renderNodes(doc.nodes, style),
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
  const content = renderNodes(nodes, style, true);

  // Word will not accept an empty part, so a strip that shows nothing still
  // gets a paragraph — a hairline one. An ordinary empty paragraph is a full
  // line deep, which is a band of white space on the page that asked for no
  // band at all.
  return content.length === 0 ? [emptyFurnitureParagraph()] : content;
}

/** The paragraph a header or footer that shows nothing is still made of. */
function emptyFurnitureParagraph(): Paragraph {
  return new Paragraph({
    spacing: { before: 0, after: 0, line: 20, lineRule: LineRuleType.EXACT },
    children: [new TextRun({ text: "", size: 2 })],
  });
}

/**
 * A run of sibling nodes, with page breaks folded into what follows them.
 *
 * A break is a property of where the next thing starts, not a thing on the
 * page — packed as its own paragraph it leaves an empty line at the foot of
 * the outgoing page. So a `pageBreak` node is held and handed to the next
 * sibling as `w:pageBreakBefore` on its first paragraph. The paragraph-of-its-
 * own form survives only where there is no such paragraph to carry it: a break
 * straight into a table, a break at the very end, two breaks in a row.
 */
function renderNodes(
  nodes: readonly DocumentNode[],
  style: DocumentStyle,
  furniture = false,
  breakBefore = false,
  cell?: CellContext,
): (Paragraph | Table)[] {
  const blocks: (Paragraph | Table)[] = [];
  let pending = breakBefore;

  for (const node of nodes) {
    if (node.kind === "pageBreak") {
      if (pending) {
        // Two breaks in a row: the first has nothing to attach to, so it keeps
        // its own paragraph and produces the blank page it asks for.
        blocks.push(pageBreakParagraph());
      }
      pending = true;
      continue;
    }

    const rendered = renderNode(node, style, furniture, pending, cell);

    if (rendered.length === 0) {
      continue;
    }

    // A table cannot carry the break itself — Word has no break-before on one
    // and docx-preview splits only on top-level elements — so a paragraph goes
    // in front to carry it. It is the separator paragraph wearing the break
    // style: an exact line of one point, which lands at the head of the new
    // page instead of leaving a full empty line at the foot of the old one.
    if (pending && rendered[0] instanceof Table) {
      blocks.push(tableSeparatorParagraph(true));
    } else if (blocks.at(-1) instanceof Table && rendered[0] instanceof Table) {
      // Two tables in a row: Word would read them as one. (The break
      // paragraph above already separates them in the other branch.)
      blocks.push(tableSeparatorParagraph());
    }
    blocks.push(...rendered);
    pending = false;
  }

  if (pending) {
    blocks.push(pageBreakParagraph());
  }

  return blocks;
}

/** The legacy form: a paragraph holding only the break. */
function pageBreakParagraph(): Paragraph {
  return new Paragraph({ children: [new PageBreak()] });
}

/**
 * The paragraph that keeps two tables from becoming one.
 *
 * Word reads two `w:tbl` elements written back to back as a single table and
 * lays the second one out on the first one's grid — so a band of dates
 * following a letterhead would take the letterhead's columns. A paragraph
 * between them is the only thing that says they are two, which is why Word
 * writes one itself whenever a document puts tables in a row.
 *
 * It is set to an exact line of one point with no spacing: it is there to
 * separate, not to take up room. Nothing before this mattered while every
 * section printed a heading between its tables — suppressing those headings
 * is what put tables next to each other for the first time.
 */
function tableSeparatorParagraph(carriesBreak = false): Paragraph {
  return new Paragraph({
    ...(carriesBreak ? { style: BREAK_STYLE_ID } : {}),
    spacing: { before: 0, after: 0, line: 20, lineRule: LineRuleType.EXACT },
    children: [new TextRun({ text: "", size: 2 })],
  });
}

/**
 * The two styles a carried break rides on.
 *
 * Word honours a `w:pageBreakBefore` written straight onto a paragraph, but
 * docx-preview — which is how a packed file is read back and shown — takes the
 * property only from a paragraph's style, never from its direct formatting. A
 * break written directly turns the page in Word and is silently dropped in the
 * preview, so one file would paginate two different ways. Riding the style is
 * the single form both engines read.
 *
 * There are two because a paragraph can only name one style: the heading a
 * section prints already names `Heading1`, so the break it carries has to come
 * from a style based on that one rather than instead of it.
 */
const BREAK_STYLE_ID = "PageBreakBefore";
const BREAK_HEADING_STYLE_ID = "PageBreakBeforeHeading1";

/** The paragraph option carrying a held break, or nothing when none is held. */
function breakStyle(broken: boolean | undefined): { style?: string } {
  return broken === true ? { style: BREAK_STYLE_ID } : {};
}

/**
 * The two break-carrying styles.
 *
 * `docx` builds a paragraph style's properties with the same
 * `ParagraphProperties` it builds a paragraph's with, so `pageBreakBefore`
 * packs correctly here — but `IParagraphStylePropertiesOptions`, the type it
 * accepts, leaves the property out. The cast is that gap and nothing more: the
 * emitted `<w:pPr><w:pageBreakBefore/></w:pPr>` is asserted in the tests.
 *
 * (`importedStyles` looks like the tidier door and is not: `options.styles`
 * spreads over the default set, so supplying it drops `docDefaults` and every
 * built-in heading style.)
 *
 * The plain style names no `basedOn`, so it inherits the document defaults
 * exactly as an unstyled paragraph does. The heading one is based on
 * `Heading1`, so a section that turns the page still looks like every other
 * section heading.
 */
function breakStyles(): IParagraphStyleOptions[] {
  const carriesBreak = { pageBreakBefore: true } as IParagraphStylePropertiesOptions;

  return [
    {
      id: BREAK_STYLE_ID,
      name: "Page Break Before",
      next: "Normal",
      paragraph: carriesBreak,
    },
    {
      id: BREAK_HEADING_STYLE_ID,
      name: "Page Break Before Heading 1",
      basedOn: "Heading1",
      next: "Normal",
      paragraph: carriesBreak,
    },
  ];
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
  return new Blob([await renderDocxBytes(doc) as BlobPart], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
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
  // Charts go in after the fact because they are parts of the package rather
  // than elements of the document, and the packer owns the zip. A document
  // with none comes back from this untouched.
  return await addChartsToPackage(
    new Uint8Array(await Packer.toBuffer(createDocxDocument(doc))),
    doc,
  );
}

/**
 * What a cell imposes on whatever it holds: the column's alignment, and the
 * leading and zero spacing that belong inside a cell rather than between
 * blocks of prose.
 */
type CellContext = {
  alignment: (typeof AlignmentType)[keyof typeof AlignmentType] | undefined;
  line: number;
  /**
   * The room either side, in twips, for a container that cannot state its own.
   *
   * A cell leaves this alone — it says its padding in `w:tcMar`, which both
   * engines read. A shape has nowhere to put it: VML states the room as an
   * `inset` on the text box, and docx-preview turns a `v:textbox` into a bare
   * `foreignObject` without ever reading that attribute, so the words sit
   * against the box's edge on screen while Word indents them properly.
   */
  indent?: { left: number; right: number };
};

/** The paragraph options a cell context contributes, or none outside a cell. */
function inCell(cell: CellContext | undefined) {
  return cell === undefined ? {} : {
    alignment: cell.alignment,
    spacing: { after: 0, line: cell.line },
    ...(cell.indent ? { indent: cell.indent } : {}),
  };
}

function renderNode(
  node: DocumentNode,
  style: DocumentStyle,
  furniture = false,
  breakBefore = false,
  cell?: CellContext,
): (Paragraph | Table)[] {
  // `breakBefore` is a held page break from `renderNodes`, landing on the
  // first paragraph this node produces.
  const broken = breakBefore || undefined;

  if (node.kind === "section") {
    // A section named `showTitle: false` keeps its title for the model and
    // prints nothing — its content already says what it is. The held break
    // then falls through to the first child instead.
    if (node.showTitle === false) {
      return renderNodes(node.children, style, furniture, breakBefore);
    }

    return [
      new Paragraph(
        broken === true
          ? { text: node.title ?? node.id, style: BREAK_HEADING_STYLE_ID }
          : { text: node.title ?? node.id, heading: HeadingLevel.HEADING_1 },
      ),
      ...renderNodes(node.children, style, furniture),
    ];
  }

  if (node.kind === "paragraph") {
    const block = blockOf(style, node.variant);

    return [
      new Paragraph({
        children: paragraphRuns(node, block, style),
        ...breakStyle(broken),
        // The node's alignment, then the theme's, then nothing at all — a
        // paragraph that never said should write no `w:jc`, so it goes on
        // inheriting whatever the style around it decides.
        alignment: textAlignmentOf(node.align ?? block?.align),
        shading: block?.fill === undefined
          ? undefined
          : { type: ShadingType.CLEAR, fill: block.fill },
        border: blockBorders(block, "paragraph"),
        indent: blockIndent(block, style),
        spacing: {
          ...blockSpacingBefore(block),
          after: ptToTwips(block?.spacingAfterPt ?? style.paragraph.spacingAfterPt),
          ...blockLine(block, style),
        },
        ...blockKeeps(block),
        ...blockTabStops(block),
      }),
    ];
  }

  if (node.kind === "image") {
    const block = blockOf(style, node.variant);
    const picture = new Paragraph({
      children: [imageRunOf(node)],
      ...breakStyle(broken),
      ...inCell(cell),
    });

    // A picture that names a variant draws it. Word has no box around a run,
    // so the box is a single-cell table — which is also what makes the card
    // hold its shape whether the picture has arrived yet or not.
    return blockDrawsABox(block) && cell === undefined
      ? [imageCard(node, block, picture)]
      : [picture];
  }

  if (node.kind === "graph") {
    return renderGraph(node, style, broken, cell);
  }

  if (node.kind === "table") {
    return [renderTable(node, style, furniture)];
  }

  if (node.kind === "shape") {
    return [renderShape(node, style, furniture)];
  }

  // A page break met head-on rather than between siblings — inside a cell, or
  // handed in alone. Only the paragraph-of-its-own form is possible here.
  if (node.kind === "pageBreak") {
    return [pageBreakParagraph()];
  }

  // Word counts the pages itself, as fields it re-evaluates whenever the
  // document is repaginated. A number written in at build time would be a
  // number that stops being true the moment an engine writes a longer
  // paragraph into a dynamic node.
  if (node.kind === "pageNumber") {
    return [new Paragraph({ children: pageNumberRuns(node), ...breakStyle(broken), ...inCell(cell) })];
  }

  // A built document has already walked its loops, so this is a published one
  // being packed directly. The body is all there is to show: one pass, standing
  // for however many the request will ask for.
  //
  // A row or a cell reaching here is one outside the table it belongs to, and
  // its content is still content — printing it beats dropping it silently.
  if (node.kind === "repeat" || node.kind === "tableRow" || node.kind === "tableCell") {
    return renderNodes(node.children, style, furniture, breakBefore, cell);
  }

  return [
    new Paragraph({
      children: [new TextRun(node.title ?? "Table of contents")],
      ...breakStyle(broken),
      ...inCell(cell),
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
 * The block something is drawn with, from the widest naming to the narrowest.
 *
 * Each statement overrides the last property by property. A cell that names a
 * variant is saying what is different about it, not offering a replacement for
 * everything its row and table already said — take it as a replacement and a
 * cell wanting the band's tint and its own leading has to restate the tint,
 * and a cell that forgets comes out untinted inside a tinted band.
 *
 * Runs already merged this way where a paragraph's variant met its cell's, so
 * this is the same rule reaching the rest of the block.
 */
function blockFor(
  style: DocumentStyle,
  ...variants: (string | undefined)[]
): DocumentBlockStyle | undefined {
  const named = variants
    .map((variant) => blockOf(style, variant))
    .filter((block): block is DocumentBlockStyle => block !== undefined);

  if (named.length === 0) {
    return undefined;
  }

  const merged: Record<string, unknown> = {};

  for (const block of named) {
    for (const [property, value] of Object.entries(block)) {
      // A property a block leaves out is one it has no opinion on, so it must
      // not overwrite the opinion underneath it.
      if (value !== undefined) {
        merged[property] = value;
      }
    }
  }

  return merged as DocumentBlockStyle;
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
  font?: string;
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
    font: block.font,
  };
}

/**
 * The runs one paragraph prints: its words, with any pictures set among them.
 *
 * The text is cut at each picture's offset so the two arrive in the order they
 * were written. A paragraph with no pictures is the single text run it always
 * was — including the empty one, which Word needs to keep the line.
 */
function paragraphRuns(
  node: ParagraphNode,
  block: DocumentBlockStyle | undefined,
  style: DocumentStyle,
): (TextRun | ImageRun)[] {
  const text = node.text ?? "";
  const inline = [...(node.inlineImages ?? [])].sort((a, b) => a.at - b.at);

  if (inline.length === 0) {
    return [textRun(text, blockRun(block, style))];
  }

  const runs: (TextRun | ImageRun)[] = [];
  let cut = 0;

  for (const { at, image } of inline) {
    const before = text.slice(cut, Math.max(cut, Math.min(at, text.length)));
    if (before.length > 0) {
      runs.push(textRun(before, blockRun(block, style)));
    }
    runs.push(imageRunOf(image));
    cut = Math.max(cut, Math.min(at, text.length));
  }

  const rest = text.slice(cut);
  if (rest.length > 0 || runs.length === 0) {
    runs.push(textRun(rest, blockRun(block, style)));
  }

  return runs;
}

/**
 * One run of text, with any tabs in it written as tabs.
 *
 * A tab character left inside `<w:t>` is not a tab: it is whitespace, and what
 * a tab stop aligns to is the `<w:tab/>` element. Word is forgiving enough to
 * draw the literal one anyway, which is worse than if it were not — the file
 * looks right in Word and collapses to a single space everywhere else,
 * including in the preview.
 *
 * So the text is cut at every tab and the element goes in between. A string
 * with no tab in it is the single run it always was.
 */
function textRun(text: string, props: Record<string, unknown>): TextRun {
  if (!text.includes("\t")) {
    return new TextRun({ text, ...props });
  }

  const pieces = text.split("\t");
  const children: (string | Tab)[] = [];

  pieces.forEach((piece, index) => {
    if (index > 0) {
      children.push(new Tab());
    }
    if (piece.length > 0) {
      children.push(piece);
    }
  });

  return new TextRun({ children, ...props });
}

/**
 * A cell paragraph's runs, with its pictures set among them.
 *
 * The cell path builds its text run itself — it has the column's alignment and
 * the cell's own block to fold in first — so the pictures are placed around
 * that finished run rather than rebuilt from the node.
 */
function withInlineImages(node: ParagraphNode, textRuns: TextRun[]): (TextRun | ImageRun)[] {
  const inline = [...(node.inlineImages ?? [])].sort((a, b) => a.at - b.at);
  if (inline.length === 0) return textRuns;

  // A picture at offset 0 leads the line — the mark before the words. Anything
  // further in follows them, because one text run cannot be cut in two here
  // without losing the styling that was just resolved onto it.
  const leading = inline.filter((entry) => entry.at === 0).map((entry) => imageRunOf(entry.image));
  const trailing = inline.filter((entry) => entry.at > 0).map((entry) => imageRunOf(entry.image));

  return [...leading, ...textRuns, ...trailing];
}

/** Whether a block draws anything around what it holds. */
function blockDrawsABox(block: DocumentBlockStyle | undefined): block is DocumentBlockStyle {
  return block !== undefined &&
    (block.border !== undefined || block.fill !== undefined || block.paddingPt !== undefined);
}

/**
 * The box a picture's variant draws around it.
 *
 * A single-cell table, because Word has no box around a run. The cell is sized
 * to the picture plus its padding, so the card holds the same shape whether
 * the picture has arrived or is still a label standing in for one — a page
 * laid out around a string and re-laid-out around a 38mm square when the real
 * one turns up is a page that moves under the reader.
 */
function imageCard(
  node: ImageNode,
  block: DocumentBlockStyle,
  picture: Paragraph,
): Table {
  const padding = block.paddingPt ?? 0;
  const side = (node.width ?? 120) + padding * 2;
  const width = Math.round(side * 20);

  return new Table({
    columnWidths: [width],
    width: { size: width, type: WidthType.DXA },
    // Sized to the picture, so it is not Word's to resize. Same reason as
    // `renderTable`: without this the card is autofit around whatever is in it.
    layout: TableLayoutType.FIXED,
    borders: block.border === undefined ? gridlessBorders : undefined,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: blockBorders(block, "cell"),
            shading: block.fill === undefined
              ? undefined
              : { type: ShadingType.CLEAR, fill: block.fill },
            margins: {
              marginUnitType: WidthType.DXA,
              top: ptToTwips(padding),
              bottom: ptToTwips(padding),
              left: ptToTwips(padding),
              right: ptToTwips(padding),
            },
            children: [picture],
          }),
        ],
      }),
    ],
  });
}

/**
 * The room a cell leaves around what it holds, side by side.
 *
 * `paddingPt` sets all four; `paddingSidesPt` names the ones that differ. A
 * theme that says neither gets the room the screen leaves, since a column that
 * lines up in the preview and not in Word is the preview lying.
 */
function cellPadding(
  block: DocumentBlockStyle | undefined,
  header: boolean,
): { top: number; right: number; bottom: number; left: number } {
  const vertical = block?.paddingPt ?? (header ? 6 : 5);
  const horizontal = block?.paddingPt ?? 8;
  const sides = block?.paddingSidesPt;

  return {
    top: sides?.top ?? vertical,
    bottom: sides?.bottom ?? vertical,
    left: sides?.left ?? horizontal,
    right: sides?.right ?? horizontal,
  };
}

/** How a block sits against the height of its cell, when it says. */
function verticalAlignOf(block: DocumentBlockStyle | undefined) {
  if (block?.valign === "center") return VerticalAlign.CENTER;
  if (block?.valign === "bottom") return VerticalAlign.BOTTOM;
  return block?.valign === "top" ? VerticalAlign.TOP : undefined;
}

/**
 * The leading a block is set on, in the twentieths of a point Word counts in.
 *
 * A block that names none is set on the body's, which is what every paragraph
 * did before a block could have an opinion.
 */
function blockLine(block: DocumentBlockStyle | undefined, style: DocumentStyle) {
  // A strip says its depth outright: it is a band of colour with no words in
  // it, so there is no leading to reason about, only a height.
  if (block?.heightPt !== undefined) {
    return { line: Math.round(block.heightPt * 20), lineRule: LineRuleType.EXACT };
  }

  const multiple = block?.lineHeight ?? style.typography.bodyLineHeight;
  const sizePt = block?.fontSizePt ?? style.typography.bodySizePt;

  // Stated in points, and said to be exact.
  //
  // A leading written as a multiple is a multiple of different things in the
  // two engines that have to agree about this document: Word multiplies the
  // font's own line height, CSS multiplies the font size, and Aptos puts about
  // 22% between those. Measured, that was a charge row 48px tall in Word and
  // 53.11px in the preview — the same content, laid out 10.6% apart, drifting
  // further down the page with every line.
  //
  // A number of points is the one form both read the same way. `exact` is what
  // makes Word honour it rather than growing the line to fit the face, which is
  // the same growth the preview does not do.
  // Measured against Word, over the same three-paragraph page, as the distance
  // between where the preview draws each paragraph and where Word does:
  //
  //   exact                       1.62mm   <- this
  //   auto                        2.29mm
  //   atLeast, reading corrected  2.35mm
  //   atLeast, as read today     20.93mm
  //
  // `atLeast` is the kinder rule in Word — it grows a line rather than clipping
  // one — but it is the rule the two engines agree about least, because Word
  // grows to the *face's* natural line height and CSS does not grow at all.
  // The conformance case `text/line-height` holds this to the numbers above.
  return {
    line: Math.round(multiple * sizePt * 20),
    lineRule: LineRuleType.EXACT,
  };
}

/**
 * The borders a block draws, as Word measures them.
 *
 * Width is in eighths of a point and the gap to the text in whole points, both
 * of which Word caps — a border it cannot draw is worse than one drawn a shade
 * thinner, so the numbers are clamped rather than passed through.
 *
 * `on` decides whether the padding is written into the border, and it is not a
 * detail: `w:space` means two different things on the two elements this
 * builds. On a `w:pBdr` it is the gap between the border and the text, and it
 * is the *only* way a paragraph holds its words off its own edge — Word has no
 * padding on a paragraph. On a `w:tcBorders` it is the gap between the cell's
 * edge and the border, and the room inside is `w:tcMar`'s job, which the cell
 * has already been given. Writing it on both put the padding in twice: Word
 * drew a 12pt-padded panel 24pt in, docx-preview drew it 12pt in, and every
 * row under it moved down by the difference.
 *
 * @param on Which element the borders are for. A cell has margins of its own;
 * a paragraph does not.
 */
function blockBorders(block: DocumentBlockStyle | undefined, on: "paragraph" | "cell") {
  if (!block?.border) {
    return undefined;
  }

  const edge = {
    style: BorderStyle.SINGLE,
    color: block.border,
    size: Math.max(1, Math.round((block.borderWidthPt ?? 1) * 8)),
    ...(on === "paragraph" ? { space: Math.min(31, Math.round(block.paddingPt ?? 0)) } : {}),
  };
  const sides = block.borderSides ?? ["top", "right", "bottom", "left"];

  return Object.fromEntries(sides.map((side) => [side, edge]));
}

/**
 * The hairline a row is separated from the next one by.
 *
 * One line under each body cell, in the palette's rule colour, and nothing
 * else: a heading already has its fill to set it apart, a row that draws its
 * own ground would be cut across by a rule at its foot, and running furniture
 * is not prose being separated into rows.
 *
 * A row, not a cell. A rule that covers the plain columns of a part-filled
 * row and stops at the filled ones is not a separator, it is a stray line.
 */
function separatorBorder(
  row: TableRowNode,
  rowIsFilled: boolean,
  furniture: boolean,
  style: DocumentStyle,
) {
  if (row.header === true || rowIsFilled || furniture) {
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
 * What a block indents by: past the margin to bleed, or inside it to sit in.
 *
 * Word measures a paragraph from the margin, so reaching outside one is a
 * negative indent — the same distance the margin is wide, which puts the block
 * flush with the edge of the sheet. Everything else here is positive and sits
 * within the text column.
 *
 * The three are exclusive by construction rather than by documentation: a
 * block that bleeds is not also inset, and a measure is a different way of
 * saying the same thing as a right indent. Ordering them settles which wins
 * without a document having to know it asked two questions at once.
 */
function blockIndent(block: DocumentBlockStyle | undefined, style: DocumentStyle) {
  if (block?.bleed === true) {
    return {
      left: -mmToTwips(style.page.margins.leftMm),
      right: -mmToTwips(style.page.margins.rightMm),
    };
  }

  // A measure is the text column narrowed from the right, not a new margin:
  // everything else on the page still stands where it stood, and the block
  // stops short of them.
  if (block?.maxWidthMm !== undefined) {
    const columnMm = (style.page.size === "A4" ? 210 : 215.9) -
      style.page.margins.leftMm - style.page.margins.rightMm;
    const spare = columnMm - block.maxWidthMm;

    return spare > 0 ? { right: mmToTwips(spare) } : undefined;
  }

  const indent: { left?: number; right?: number; firstLine?: number; hanging?: number } = {};

  if (block?.indentMm !== undefined) {
    indent.left = mmToTwips(block.indentMm);
  }
  if (block?.indentRightMm !== undefined) {
    indent.right = mmToTwips(block.indentRightMm);
  }

  // A hang and a first-line indent are the same attribute pulling opposite
  // ways, and Word writes only one of them. The hang wins because it is the
  // structural one: a block that hangs is shaped that way, where a first-line
  // indent is a typographic flourish on top.
  if (block?.hangingIndentMm !== undefined) {
    indent.hanging = mmToTwips(block.hangingIndentMm);
  } else if (block?.firstLineIndentMm !== undefined) {
    indent.firstLine = mmToTwips(block.firstLineIndentMm);
  }

  return Object.keys(indent).length === 0 ? undefined : indent;
}

/**
 * How a paragraph's lines sit in the column, as Word names it.
 *
 * `justify` is `both` in OOXML — both edges flush — which is a better name for
 * what happens and a worse one for what people call it. The document says the
 * word people say and the translation stops here.
 */
function textAlignmentOf(align: TextAlign | undefined) {
  if (align === undefined) {
    return undefined;
  }

  return {
    left: AlignmentType.LEFT,
    center: AlignmentType.CENTER,
    right: AlignmentType.RIGHT,
    justify: AlignmentType.JUSTIFIED,
  }[align];
}

/** The space a block leaves above itself, when it leaves any. */
function blockSpacingBefore(block: DocumentBlockStyle | undefined) {
  return block?.spacingBeforePt === undefined
    ? {}
    : { before: ptToTwips(block.spacingBeforePt) };
}

/**
 * The breaks a block refuses to be on the wrong side of.
 *
 * Only written when the block asks. A paragraph that said nothing must write
 * nothing: `w:keepNext w:val="0"` is Word being told to *allow* the break,
 * which is a different statement from having no opinion, and it would override
 * a style that had one.
 */
function blockKeeps(block: DocumentBlockStyle | undefined) {
  return {
    ...(block?.keepWithNext === true ? { keepNext: true } : {}),
    ...(block?.keepLines === true ? { keepLines: true } : {}),
  };
}

/** The stops a block's tabs align to, in the twips Word measures them in. */
function blockTabStops(block: DocumentBlockStyle | undefined) {
  if (block?.tabStopsMm === undefined || block.tabStopsMm.length === 0) {
    return {};
  }

  const types = {
    left: TabStopType.LEFT,
    center: TabStopType.CENTER,
    right: TabStopType.RIGHT,
    decimal: TabStopType.DECIMAL,
  };
  const leaders = {
    none: LeaderType.NONE,
    dot: LeaderType.DOT,
    dash: LeaderType.HYPHEN,
    underscore: LeaderType.UNDERSCORE,
  };

  return {
    tabStops: block.tabStopsMm.map((stop) => ({
      type: types[stop.align ?? "left"],
      position: mmToTwips(stop.at),
      leader: leaders[stop.leader ?? "none"],
    })),
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
  // `bleed` reaches a table the same way it reaches a paragraph: out past the
  // margins to the paper's edge. A footer bar that stopped at the margin would
  // be a bar with a white gutter either side of it, which is not a bar.
  const block = blockOf(style, node.variant);
  const bleed = block?.bleed === true;
  const widths = columnWidths(node.columns, style, bleed);
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
    // A table that does not say which layout algorithm it means is autofit,
    // and an autofit table's grid is a suggestion: Word reworks the columns
    // around the words in the cells. Measured on a 60/40/70mm table, it
    // printed 67.5/50.9/51.6 — the total right, every boundary wrong, which is
    // the shape of the bug nobody notices until a figure fails to line up. The
    // widths were worked out here precisely so a money column comes out the
    // width it was given, and `fixed` is what makes Word honour them.
    // `tables/column-widths` in the conformance suite holds this.
    layout: TableLayoutType.FIXED,
    // Word indents a table from the text column's left edge, so reaching the
    // paper's edge means indenting back by the margin.
    indent: bleed
      ? { size: -mmToTwips(style.page.margins.leftMm), type: WidthType.DXA }
      : undefined,
    // Word's own default is a black grid around every cell, which no theme
    // ever asked for and the screen never draws. The lines a table shows are
    // the ones decided below, one per row and in the palette's rule colour.
    borders: gridlessBorders,
    // A zebra is a reading aid for a column of like rows, and what says a
    // table is one of those is its header. A table without one is a layout —
    // a letterhead, a totals block, a running strip — where striping tints
    // whichever cells happen to be on an odd row. On the invoice that put a
    // grey band in the empty spacer beside the totals panel: a stripe of
    // nothing, next to the figures rather than under them.
    rows: rows.map((row, index) =>
      renderRow(row, node, style, index < headers, furniture, headers === 0 ? -1 : index - headers)),
  });
}


/**
 * The default depth of a shape nothing has said a height for, in points.
 *
 * A shape does not grow to fit its words, so something has to decide. Three
 * lines of body text at the default leading is the size a callout wants and
 * the size that makes it obvious the number is a decision rather than a fit.
 */
const SHAPE_HEIGHT_PT = 48;

/**
 * A shape: a drawn rectangle with the document's own paragraphs on top of it.
 *
 * Written as VML — `w:pict` holding a `v:rect` holding a `v:textbox` — rather
 * than as the DrawingML `wps:wsp` Word writes today, and that is a decision
 * worth the paragraph it takes to explain.
 *
 * Word draws both, identically: it reads the VML back as a real `Rectangle`
 * shape, with the width, the height, the fill and the text this gives it. The
 * difference is on the other side. docx-preview renders VML — a `v:rect`
 * becomes an SVG `<rect>` and a `v:textbox` becomes a `<foreignObject>` — and
 * has no reading of `wps:wsp` at all, so a DrawingML shape draws nothing on
 * screen. Packing the form that both engines read is what keeps the preview
 * showing the document rather than an opinion about it; the alternative was to
 * pack DrawingML and then *build* the box in the preview, which is a second
 * renderer and the thing this package exists not to have.
 *
 * One repair is still needed on the preview side, and it is a structural one
 * rather than a coat of CSS: docx-preview nests the `<foreignObject>` inside
 * the `<rect>`, where SVG will not paint it. `settleDocxPreview` moves it out
 * beside the rect. Both elements come from the file and Word draws both; only
 * their nesting is wrong.
 */
/**
 * How deep a chart is drawn against its width when nothing says otherwise.
 *
 * Seven twelfths, which is the 6in by 3.5in frame Word gives a chart it
 * inserts itself. A chart is a picture of a proportion, and a frame the reader
 * already knows the shape of is one fewer thing for them to take in.
 */
const CHART_ASPECT = 7 / 12;

/** A point, in the English Metric Units a drawing states its size in. */
const EMU_PER_PT = 12700;

/**
 * A chart, as the drawing that holds its place plus anything printed under it.
 *
 * The drawing is a frame and a relationship id, and the id is a token until the
 * package is opened — see {@linkcode ../render/docx_charts.ts | docx_charts}
 * for why it cannot be allocated here. A chart with no data yet draws its
 * placeholder instead, which is the same thing a dynamic image does: the page
 * has to be laid out before the engine has written anything, and a document
 * that showed nothing where a chart will go is a document laid out for a
 * different page than the one that will print.
 */
function renderGraph(
  node: GraphNode,
  style: DocumentStyle,
  broken: true | undefined,
  cell: CellContext | undefined,
): Paragraph[] {
  const widthPt = node.width ?? columnWidthPt(style);
  const heightPt = node.height ?? widthPt * CHART_ASPECT;

  const plot = hasChartData(node)
    ? new Paragraph({
      children: [chartDrawing(node, widthPt, heightPt) as unknown as TextRun],
      ...breakStyle(broken),
      ...inCell(cell),
    })
    : new Paragraph({
      children: [
        new TextRun(`[${node.graphType} chart: ${node.placeholder ?? node.title ?? node.id}]`),
      ],
      ...breakStyle(broken),
      ...inCell(cell),
    });

  if (node.caption === undefined) {
    return [plot];
  }

  // The caption is a paragraph of its own rather than part of the drawing:
  // Word cannot put text inside a chart frame, and a caption drawn into the
  // chart would be a caption that does not reflow, cannot be read out and does
  // not come back when the document is read.
  return [plot, chartCaption(node.caption, style, cell)];
}

/**
 * The drawing that reserves a chart's frame and names the part to draw in it.
 *
 * Built element by element rather than parsed from a string: `fromXmlString`
 * returns a component wrapping what it parsed rather than the element itself,
 * so its `w:r` came out inside a tag named `undefined` — a document Word
 * refuses to open at all rather than one that draws wrongly.
 *
 * `a` and `c` are declared here rather than on the document, which is where
 * Word itself declares them: they are used by this drawing and nothing else in
 * `document.xml`, and a package that only sometimes has charts in it should
 * not sometimes have different namespaces on its root.
 */
function chartDrawing(node: GraphNode, widthPt: number, heightPt: number): ImportedXmlComponent {
  const inline = new ImportedXmlComponent("wp:inline", {
    distT: "0",
    distB: "0",
    distL: "0",
    distR: "0",
  });

  inline.push(
    new ImportedXmlComponent("wp:extent", {
      cx: String(Math.round(widthPt * EMU_PER_PT)),
      cy: String(Math.round(heightPt * EMU_PER_PT)),
    }),
  );
  inline.push(new ImportedXmlComponent("wp:effectExtent", { l: "0", t: "0", r: "0", b: "0" }));
  inline.push(
    new ImportedXmlComponent("wp:docPr", {
      id: String(drawingId(node.id)),
      name: `Chart ${node.id}`,
      // What a screen reader reads out where a sighted reader sees the plot.
      descr: node.title ?? node.caption ?? node.id,
    }),
  );
  inline.push(new ImportedXmlComponent("wp:cNvGraphicFramePr"));

  const graphic = new ImportedXmlComponent("a:graphic", {
    "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
  });
  const graphicData = new ImportedXmlComponent("a:graphicData", {
    uri: "http://schemas.openxmlformats.org/drawingml/2006/chart",
  });

  graphicData.push(
    new ImportedXmlComponent("c:chart", {
      "xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "r:id": chartRelationshipToken(node.id),
    }),
  );
  graphic.push(graphicData);
  inline.push(graphic);

  const drawing = new ImportedXmlComponent("w:drawing");
  drawing.push(inline);

  // `w:drawing` belongs inside a run, the same way `w:pict` does — see
  // `renderShape` for what a misplaced one costs.
  const run = new ImportedXmlComponent("w:r");
  run.push(drawing);

  return run;
}

/**
 * The number a drawing is known by inside the document.
 *
 * Derived from the node's id rather than counted, so that packing the same
 * document twice produces the same file. Two charts collide only if their ids
 * hash together, and a collision costs nothing a reader sees — Word takes the
 * ids as a hint, and the library packing everything else here already gives
 * every picture the number one.
 */
function drawingId(id: string): number {
  let hash = 0x811c9dc5;

  for (let index = 0; index < id.length; index += 1) {
    hash = Math.imul(hash ^ id.charCodeAt(index), 0x01000193) >>> 0;
  }

  return 100000 + (hash % 1000000);
}

/** The line printed under a chart, set back from the body it explains. */
function chartCaption(
  caption: string,
  style: DocumentStyle,
  cell: CellContext | undefined,
): Paragraph {
  const block = blockOf(style, "chartCaption");

  return new Paragraph({
    children: [
      new TextRun({
        text: caption,
        // A caption is not body prose and should not read as it: a point
        // smaller and in the ink the theme sets its secondary text in. A theme
        // that names a `chartCaption` block says it properly and wins.
        size: Math.round((block?.fontSizePt ?? style.typography.bodySizePt - 1) * 2),
        color: block?.color ?? style.palette?.muted,
        font: block?.font ?? style.typography.bodyFont,
      }),
    ],
    alignment: textAlignmentOf(block?.align ?? "center"),
    spacing: {
      after: ptToTwips(block?.spacingAfterPt ?? style.paragraph.spacingAfterPt),
    },
    ...inCell(cell),
  });
}

function renderShape(
  node: ShapeNode,
  style: DocumentStyle,
  furniture: boolean,
): Paragraph {
  const block = blockOf(style, node.variant);
  const inset = cellPadding(block, false);
  const widthPt = node.width ?? columnWidthPt(style);
  const heightPt = node.height ?? block?.heightPt ?? SHAPE_HEIGHT_PT;

  const rect = new ImportedXmlComponent("v:rect", {
    // VML takes its geometry from a CSS-shaped style string, in points. This
    // is the same number the preview lays the `<svg>` out at, because it is
    // the same attribute read by both.
    style: `width:${round2(widthPt)}pt;height:${round2(heightPt)}pt`,
    fillcolor: `#${block?.fill ?? style.palette?.accent ?? "1F2933"}`,
    ...(block?.border === undefined
      ? { stroked: "f" }
      : {
        strokecolor: `#${block.border}`,
        strokeweight: `${round2(block.borderWidthPt ?? 1)}pt`,
      }),
  });

  // The same rule, said again as a child element.
  //
  // VML states a stroke twice over: as attributes on the shape, and as a
  // `v:stroke` child. Word reads either. docx-preview reads only the child —
  // its `parseStroke` takes `color` and `weight` off a `<v:stroke>` and never
  // looks at `strokecolor` — so a shape ruled only by the attributes is ruled
  // in Word and bare on screen. Writing both is what Word itself writes, and
  // they cannot disagree because they are built from the one block.
  if (block?.border !== undefined) {
    rect.push(
      new ImportedXmlComponent("v:stroke", {
        color: `#${block.border}`,
        weight: `${round2(block.borderWidthPt ?? 1)}pt`,
      }),
    );
  }

  const textbox = new ImportedXmlComponent("v:textbox", {
    // The room inside the box, in the order VML states it: left, top, right,
    // bottom. The same `paddingPt` a cell reads, so a variant means one thing
    // whichever kind of box the theme draws it in.
    //
    // Only the vertical half is stated here, and that is the repair rather
    // than an omission. docx-preview renders a `v:textbox` as a bare
    // `foreignObject` and never reads `inset`, so room stated only here is
    // room Word leaves and the screen does not — the words sit hard against
    // the fill in the preview and indented in the file, which is the preview
    // lying about the document. The sides go on the paragraphs instead, as
    // `w:ind`, which both engines read; stating them in both places would
    // simply have Word leave the gap twice.
    inset: `0pt,${round2(inset.top)}pt,0pt,${round2(inset.bottom)}pt`,
  });

  const content = new ImportedXmlComponent("w:txbxContent");

  for (const child of shapeContent(node, block, style, furniture, inset, heightPt)) {
    content.push(child);
  }

  textbox.push(content);
  rect.push(textbox);

  // `w:pict` belongs inside a run, not beside one. Word repairs a misplaced
  // one silently and docx-preview does not — it looks for a picture among a
  // run's children and finds nothing, so the shape draws in Word and vanishes
  // on screen.
  const run = new ImportedXmlComponent("w:r");
  const pict = new ImportedXmlComponent("w:pict");

  pict.push(rect);
  run.push(pict);

  return new Paragraph({
    alignment: textAlignmentOf(block?.align),
    spacing: {
      ...blockSpacingBefore(block),
      after: ptToTwips(block?.spacingAfterPt ?? style.paragraph.spacingAfterPt),
    },
    children: [run as unknown as TextRun],
  });
}

/**
 * What is drawn on the box.
 *
 * A cell's rules, not a body paragraph's: the spacing that belongs between
 * blocks of prose does not belong inside a box, and the block's own leading
 * decides the rest. Word will not accept an empty text box, so a shape with
 * nothing on it still gets a paragraph — an empty band is a band, not a
 * missing one.
 */
function shapeContent(
  node: ShapeNode,
  block: DocumentBlockStyle | undefined,
  style: DocumentStyle,
  furniture: boolean,
  inset: { left: number; right: number },
  heightPt: number,
): Paragraph[] {
  // The room either side, carried on the paragraphs because the text box's own
  // `inset` reaches Word and not the preview. See `renderShape`.
  const indent = { left: ptToTwips(inset.left), right: ptToTwips(inset.right) };

  // Leading struck from the box actually drawn, not from the theme's fallback
  // depth. Exact leading the height of the box is what sits one line in the
  // middle of it, so a node that states its own height has to be the number
  // used — `blockLine` would reach for the variant's `heightPt` and centre the
  // words in a box of the wrong size, which on a 56pt stamp over a 44pt block
  // style left them 6pt high.
  const boxLine = { line: Math.round(heightPt * 20), lineRule: LineRuleType.EXACT };

  const content = node.children.flatMap((child): Paragraph[] => {
    if (child.kind !== "paragraph") {
      // Everything else is still a child of this shape, so it takes the
      // block's alignment and leading like the paragraphs beside it.
      return renderNode(child, style, furniture, false, {
        alignment: alignments[shapeAlign(block)],
        ...boxLine,
        indent,
      }).filter((rendered): rendered is Paragraph => rendered instanceof Paragraph);
    }

    // The shape's block, then the paragraph's own — a muted note under a
    // heading on the same box is set by its own variant, not the shape's.
    const inner = blockOf(style, child.variant);
    const run = { ...blockRun(block, style), ...blockRun(inner, style) };

    return [
      new Paragraph({
        alignment: alignments[shapeAlign(inner ?? block)],
        // A paragraph that names its own variant is set by it — a muted note
        // under a heading on the same box. One that names none is centred on
        // the box, which is what a banner's single line wants.
        spacing: { after: 0, ...(inner ? blockLine(inner, style) : boxLine) },
        indent,
        children: withInlineImages(child, [
          new TextRun({
            text: child.text ?? "",
            bold: run.bold,
            color: run.color,
            size: run.size,
            font: run.font,
            allCaps: run.allCaps,
            characterSpacing: run.characterSpacing,
          }),
        ]),
      }),
    ];
  });

  return content.length === 0 ? [new Paragraph({})] : content;
}

/** How the words sit across the box. Left unless the theme says otherwise. */
function shapeAlign(block: DocumentBlockStyle | undefined): TableAlign {
  const align = block?.align;
  return align === "center" || align === "right" ? align : "left";
}

/** The text column's width, in points — what a shape fills unless it is sized. */
function columnWidthPt(style: DocumentStyle): number {
  const pageMm = style.page.size === "A4" ? 210 : 215.9;
  const columnMm = pageMm - style.page.margins.leftMm - style.page.margins.rightMm;

  return columnMm * (72 / 25.4);
}

/** Two decimal places, which is finer than Word or a screen can draw. */
function round2(value: number): number {
  return Math.round(value * 100) / 100;
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
  bodyIndex = -1,
): TableRow {
  // F8: striping is the renderer's to do, counted as the rows are drawn. A
  // variant chosen by a `.map` index works in preview and dies on publish —
  // the map becomes one loop the engine walks, and a variant is a static
  // string, so every row would get whatever the build happened to decide.
  //
  // `bodyIndex` is -1 for every row of a table with no header, which is how
  // a layout table opts out without having to say so.
  const alt = blockOf(style, "rowAlt");
  const striped = alt !== undefined && bodyIndex >= 0 && bodyIndex % 2 === 1;
  const cells: TableCell[] = [];
  let column = 0;

  // Whether any of this row draws its own ground, worked out before a single
  // cell is built. The separator is a property of the row — deciding it from
  // each cell's own fill drew it under the columns that happened to be plain
  // and not under the ones that were not, which on a totals block is a rule
  // to the left of the panel that stops where the panel starts.
  const rowIsFilled = row.children.some((child) =>
    child.kind === "tableCell" &&
    (blockFor(style, table.variant, row.variant, child.variant)?.fill !== undefined ||
      (striped && alt?.fill !== undefined))
  );

  for (const child of row.children) {
    if (child.kind !== "tableCell") {
      continue;
    }

    const span = child.span ?? 1;
    const align = child.align ?? table.columns[column]?.align ?? "left";

    // The table's block, then its row's, then the cell's — each overriding the
    // last property by property rather than wholesale.
    const block = blockFor(style, table.variant, row.variant, child.variant);
    // The accent bar over a heading is what a header row looks like when the
    // theme has not named one. Once it has — for the cell, its row or its
    // table — that naming is the whole answer, and the default steps aside
    // rather than showing through the half of it the theme left unsaid.
    const fill = block?.fill ??
      (row.header && block === undefined ? style.palette?.accent ?? "1F2933" : undefined) ??
      (striped ? alt?.fill : undefined);
    // The theme's padding, or the room a cell is given when it says nothing —
    // the same room the screen leaves, since a column that lines up in the
    // preview and not in Word is the preview lying about the document.
    const inset = cellPadding(block, row.header === true);

    cells.push(
      new TableCell({
        columnSpan: span > 1 ? span : undefined,
        verticalAlign: verticalAlignOf(block),
        shading: fill === undefined ? undefined : { type: ShadingType.CLEAR, fill },
        borders: blockBorders(block, "cell") ??
          (block?.borderSides?.length === 0
            // Naming no edges is a decision, not an omission: it says this
            // block draws none, so the row hairline does not step in.
            ? undefined
            : separatorBorder(row, rowIsFilled, furniture, style)),
        margins: {
          marginUnitType: WidthType.DXA,
          top: ptToTwips(inset.top),
          bottom: ptToTwips(inset.bottom),
          left: ptToTwips(inset.left),
          right: ptToTwips(inset.right),
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
      // Everything else — a page number, a table of contents, a graph — is
      // still a child of this cell, so it takes the column's alignment and the
      // cell's spacing like the paragraphs beside it. Rendered without them it
      // prints left in a right-aligned column and carries the 6pt gap that
      // belongs between blocks of prose, which is most of why a one-line
      // footer bar came out three lines deep.
      return renderNode(child, style, furniture, false, {
        alignment: alignments[align],
        ...blockLine(block, style),
      });
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
        // The paragraph's own variant decides the leading where it has one —
        // a muted note under a description is set tighter than the row it
        // sits in — and the cell's block decides it otherwise.
        spacing: { after: 0, ...blockLine(inner ?? block, style) },
        children: withInlineImages(child, [
          new TextRun({
            text: child.text ?? "",
            bold: run.bold ?? heading.bold,
            color: run.color ?? heading.color,
            size,
            font: run.font,
            allCaps: run.allCaps ?? heading.allCaps,
            characterSpacing: run.characterSpacing ??
              (heading.trackedEm === undefined
                ? undefined
                : trackingOf(heading.trackedEm, size, style)),
          }),
        ]),
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
/**
 * Letter spacing in ems, as the twentieths of a point Word counts in, for a
 * block whose size is already known in points.
 */
function trackingPt(em: number | undefined, sizePt: number): number | undefined {
  return em === undefined ? undefined : Math.round(em * sizePt * 20);
}

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
function columnWidths(
  columns: readonly TableColumn[],
  style: DocumentStyle,
  bleed = false,
): number[] {
  // A bleeding table is laid out on the paper, not on the text column: it is
  // pulled out to the edges by a negative indent, so the width it has to fill
  // is the whole sheet.
  const available = bleed ? pageWidthTwips(style) : pageWidthTwips(style) -
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
            ...blockLine(undefined, style),
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
          characterSpacing: trackingPt(style.title.letterSpacingEm, style.title.fontSizePt),
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
          characterSpacing: trackingPt(
            style.sectionHeading.letterSpacingEm,
            style.sectionHeading.fontSizePt,
          ),
        },
        paragraph: {
          spacing: {
            before: ptToTwips(style.sectionHeading.spacingBeforePt),
            after: ptToTwips(style.sectionHeading.spacingAfterPt),
          },
        },
      },
    },
    paragraphStyles: breakStyles(),
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
