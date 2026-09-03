/**
 * Probe A — what the packed file actually says.
 *
 * A `.docx` is a zip of XML, so the only way to know a property survived
 * packing is to open the file and read it back. This probe is paragraph-shaped
 * on purpose: it walks `word/document.xml` and emits one record per `<w:p>`,
 * with the paragraph properties resolved and the first run's properties beside
 * them. Everything a paragraph case asks about — alignment, indent, spacing,
 * keeps, tabs, shading, borders — is a field on that record.
 *
 * Nothing here throws on a missing element. A property that was never written
 * comes back `null`, which is a measurement: it is exactly the finding a case
 * about an unsupported feature exists to record.
 *
 * @module
 */

import { inflateRawSync } from "node:zlib";

/** A twip is a twentieth of a point; Word counts almost everything in them. */
export const TWIPS_PER_PT = 20;
/** A millimetre, in twips. */
export const TWIPS_PER_MM = (72 / 25.4) * TWIPS_PER_PT;

/**
 * One entry out of a zip, via its central directory.
 *
 * Ported from `tests/docx.ts` rather than shared with it: the tests must keep
 * working with no dependency on the conformance package, and sixty lines of
 * zip reading is a cheaper duplicate than a coupling between them.
 */
export function entryOf(zip, wanted) {
  const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  const text = new TextDecoder("latin1").decode(zip);
  const end = text.lastIndexOf("PK");

  if (end === -1) {
    throw new Error("not a zip");
  }

  let at = view.getUint32(end + 16, true);
  const count = view.getUint16(end + 10, true);

  for (let index = 0; index < count; index += 1) {
    const nameLength = view.getUint16(at + 28, true);
    const extraLength = view.getUint16(at + 30, true);
    const commentLength = view.getUint16(at + 32, true);
    const name = text.slice(at + 46, at + 46 + nameLength);

    if (name === wanted) {
      const method = view.getUint16(at + 10, true);
      const size = view.getUint32(at + 20, true);
      const offset = view.getUint32(at + 42, true);
      const localName = view.getUint16(offset + 26, true);
      const localExtra = view.getUint16(offset + 28, true);
      const start = offset + 30 + localName + localExtra;
      const data = zip.slice(start, start + size);

      return method === 0 ? data : inflateRawSync(data);
    }

    at += 46 + nameLength + extraLength + commentLength;
  }

  throw new Error(`no ${wanted} in the package`);
}

/** Every part the package holds, by name. */
export function partNames(zip) {
  const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  const text = new TextDecoder("latin1").decode(zip);
  const end = text.lastIndexOf("PK");
  const count = view.getUint16(end + 10, true);
  const names = [];

  let at = view.getUint32(end + 16, true);

  for (let index = 0; index < count; index += 1) {
    const nameLength = view.getUint16(at + 28, true);
    const extraLength = view.getUint16(at + 30, true);
    const commentLength = view.getUint16(at + 32, true);

    names.push(text.slice(at + 46, at + 46 + nameLength));
    at += 46 + nameLength + extraLength + commentLength;
  }

  return names;
}

/** One part of the package, as text. */
export function partText(zip, part) {
  return new TextDecoder().decode(entryOf(zip, part));
}

// ---------------------------------------------------------------------------
// Reading the XML
//
// Regex over string, not a parser. The shapes here are shallow and fixed —
// `w:pPr` holds flat children, a run holds one `w:rPr` — and a dependency-free
// probe is one that keeps running when the parser it would have used moves on.
// ---------------------------------------------------------------------------

/** The value of one attribute on an element, or null. */
function attr(xml, name) {
  const match = new RegExp(`${name}="([^"]*)"`).exec(xml ?? "");
  return match === null ? null : match[1];
}

/** The first `<name …/>` or `<name …>…</name>`, whole. */
function element(xml, name) {
  const selfClosing = new RegExp(`<${name}(\\s[^>]*)?/>`).exec(xml ?? "");
  const paired = new RegExp(`<${name}(\\s[^>]*)?>([\\s\\S]*?)</${name}>`).exec(xml ?? "");

  if (paired !== null && (selfClosing === null || paired.index < selfClosing.index)) {
    return paired[0];
  }
  return selfClosing === null ? null : selfClosing[0];
}

/** Every `<name …/>` or `<name …>…</name>` in order. */
function elements(xml, name) {
  const pattern = new RegExp(`<${name}(?:\\s[^>]*)?(?:/>|>[\\s\\S]*?</${name}>)`, "g");
  return [...(xml ?? "").matchAll(pattern)].map((match) => match[0]);
}

/** An attribute of a child element, or null when either is missing. */
function childAttr(xml, name, attribute) {
  const child = element(xml, name);
  return child === null ? null : attr(child, attribute);
}

/** The same, as a number. */
function childNum(xml, name, attribute) {
  const value = childAttr(xml, name, attribute);
  if (value === null) return null;
  const number = Number.parseFloat(value);
  return Number.isFinite(number) ? number : null;
}

/**
 * A boolean toggle property, in the three states OOXML gives it.
 *
 * `<w:keepNext/>` is on, `<w:keepNext w:val="0"/>` is explicitly off, and no
 * element at all is `null` — inherit. The three are different facts, and a
 * probe that collapsed the last two would report a document as saying
 * something it never said.
 */
function toggle(xml, name) {
  const found = element(xml, name);
  if (found === null) return null;
  const value = attr(found, "w:val");
  if (value === null) return true;
  return !["0", "false", "off"].includes(value);
}

/** The text a run of XML prints. */
export function textOf(xml) {
  return [...(xml ?? "").matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)]
    .map((match) => decodeEntities(match[1]))
    .join("");
}

function decodeEntities(value) {
  return value
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&#39;", "'")
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

/** The border on one edge of a `w:pBdr`, or null when that edge is undrawn. */
function borderOf(pBdr, edge) {
  if (pBdr === null) return null;
  const found = element(pBdr, `w:${edge}`);
  if (found === null) return null;

  const style = attr(found, "w:val");
  if (style === null || style === "none" || style === "nil") return null;

  return {
    style,
    color: attr(found, "w:color"),
    sizeEighthPt: Number.parseFloat(attr(found, "w:sz") ?? "0") || null,
    spacePt: Number.parseFloat(attr(found, "w:space") ?? "0") || null,
  };
}

/**
 * One paragraph, read whole.
 *
 * The shape is deliberately flat: a case asserts `para("Centred").jc`, not a
 * path through three levels of XML. Every field is the OOXML fact under its
 * OOXML name, so an assertion can be checked against the spec without a
 * translation table in between.
 */
function readParagraph(xml, index) {
  const pPr = element(xml, "w:pPr");
  const rPr = element(element(xml, "w:r") ?? "", "w:rPr");
  const spacing = pPr === null ? null : element(pPr, "w:spacing");
  const ind = pPr === null ? null : element(pPr, "w:ind");
  const pBdr = pPr === null ? null : element(pPr, "w:pBdr");
  const tabs = pPr === null ? null : element(pPr, "w:tabs");

  return {
    index,
    text: textOf(xml),

    /** The named style, when the paragraph takes one. */
    style: childAttr(pPr, "w:pStyle", "w:val"),

    // --- the properties a paragraph case is about ---------------------------

    /** Alignment: left / center / right / both (justified) / null. */
    jc: childAttr(pPr, "w:jc", "w:val"),

    /** Indents, in twips. `firstLine` and `hanging` are mutually exclusive. */
    ind: ind === null ? null : {
      left: numAttr(ind, "w:left") ?? numAttr(ind, "w:start"),
      right: numAttr(ind, "w:right") ?? numAttr(ind, "w:end"),
      firstLine: numAttr(ind, "w:firstLine"),
      hanging: numAttr(ind, "w:hanging"),
    },

    /** Spacing, in twips, plus how the line rule is meant. */
    spacing: spacing === null ? null : {
      before: numAttr(spacing, "w:before"),
      after: numAttr(spacing, "w:after"),
      line: numAttr(spacing, "w:line"),
      lineRule: attr(spacing, "w:lineRule"),
    },

    /** The three-state keeps. */
    keepNext: toggle(pPr, "w:keepNext"),
    keepLines: toggle(pPr, "w:keepLines"),
    widowControl: toggle(pPr, "w:widowControl"),
    pageBreakBefore: toggle(pPr, "w:pageBreakBefore"),

    /** Every tab stop the paragraph declares, in order. */
    tabs: tabs === null ? [] : elements(tabs, "w:tab").map((stop) => ({
      pos: Number.parseFloat(attr(stop, "w:pos") ?? "0") || 0,
      val: attr(stop, "w:val"),
      leader: attr(stop, "w:leader"),
    })),

    /** Paragraph shading, as the fill hex without a `#`. */
    shd: childAttr(pPr, "w:shd", "w:fill"),

    /** Borders, by edge. An undrawn edge is null. */
    pBdr: pBdr === null ? null : {
      top: borderOf(pBdr, "top"),
      right: borderOf(pBdr, "right"),
      bottom: borderOf(pBdr, "bottom"),
      left: borderOf(pBdr, "left"),
    },

    // --- the first run, for the properties a paragraph sets through it ------

    run: {
      font: childAttr(rPr, "w:rFonts", "w:ascii"),
      /** Half-points, as OOXML counts them. */
      szHalfPt: childNum(rPr, "w:sz", "w:val"),
      color: childAttr(rPr, "w:color", "w:val"),
      bold: toggle(rPr, "w:b"),
      italics: toggle(rPr, "w:i"),
      caps: toggle(rPr, "w:caps"),
      /** Letter spacing, in twips. */
      spacing: childNum(rPr, "w:spacing", "w:val"),
    },

    /** Whether the paragraph holds a picture. */
    hasDrawing: xml.includes("<w:drawing"),
    /** Whether the paragraph holds a field — PAGE, NUMPAGES, TOC. */
    hasField: xml.includes("<w:fldChar") || xml.includes("<w:instrText"),
    /** Whether the paragraph carries an explicit break. */
    hasBreak: xml.includes("<w:br"),

    /** The paragraph's own XML, for the assertions no field anticipates. */
    xml,
  };
}

function numAttr(xml, name) {
  const value = attr(xml, name);
  if (value === null) return null;
  const number = Number.parseFloat(value);
  return Number.isFinite(number) ? number : null;
}

// ---------------------------------------------------------------------------
// Tables
//
// A table is the one shape in this file that nests: a cell holds paragraphs,
// and it may hold another table. So none of the reading below can use the
// first-match helpers above without saying where to stop — `element(cellXml,
// "w:tcPr")` on a plain cell that happens to contain a nested table returns
// the *inner* cell's properties, which is a measurement of the wrong cell that
// looks exactly like a measurement of the right one.
//
// Two rules keep that from happening. Children are found by a depth-aware
// scan rather than by a non-greedy regex, and properties are read only from
// the head of an element — the part before its first child — because OOXML
// always writes `w:tblPr`, `w:trPr` and `w:tcPr` first.
// ---------------------------------------------------------------------------

/**
 * Every `<name>` element directly inside this XML, ignoring any nested inside
 * one of `notInside`.
 *
 * A non-greedy `<w:tbl>[\s\S]*?</w:tbl>` stops at the first closing tag, which
 * for a table holding a table is halfway through the outer one. Depth is the
 * only way to read a shape that can contain itself.
 */
function scanElements(xml, name, notInside = []) {
  const source = xml ?? "";
  const guards = new Set(notInside);
  const found = [];

  let guardDepth = 0;
  let depth = 0;
  let start = -1;

  for (const match of source.matchAll(/<(\/?)(w:[A-Za-z]+)(?:\s[^>]*?)?(\/?)>/g)) {
    const [whole, close, tag, selfClose] = match;

    if (guards.has(tag)) {
      if (selfClose !== "/") guardDepth += close === "/" ? -1 : 1;
      continue;
    }
    if (tag !== name) continue;

    if (selfClose === "/") {
      // An empty element, written closed. It still occupies its place.
      if (guardDepth === 0 && depth === 0) found.push(whole);
      continue;
    }

    if (close === "/") {
      depth -= 1;
      if (depth === 0 && guardDepth === 0 && start !== -1) {
        found.push(source.slice(start, match.index + whole.length));
        start = -1;
      }
      continue;
    }

    if (guardDepth === 0 && depth === 0 && start === -1) start = match.index;
    depth += 1;
  }

  return found;
}

/**
 * The head of an element: everything before the first child of these names.
 *
 * This is the region an element's own properties live in, and reading them
 * from anywhere else is how a nested table's cell ends up answering for the
 * cell that holds it.
 */
function headOf(xml, ...stops) {
  const source = xml ?? "";
  const cut = stops.map((stop) => source.indexOf(`<${stop}`)).filter((index) => index !== -1);

  return cut.length === 0 ? source : source.slice(0, Math.min(...cut));
}

/**
 * What is inside an element, with its own opening and closing tags removed.
 *
 * Children are scanned for inside this, never inside the whole element. A
 * table's rows are found by skipping anything within a nested `w:tbl` — and
 * the table's own `<w:tbl>` is a `w:tbl`, so scanning the element whole finds
 * every row nested inside itself and returns none of them.
 */
function innerOf(xml) {
  const open = (xml ?? "").indexOf(">");
  const close = (xml ?? "").lastIndexOf("</");

  return open === -1 || close <= open ? "" : xml.slice(open + 1, close);
}

/** A `w:tblW`-shaped width: the number, and the unit it is counted in. */
function widthOf(xml, name) {
  const found = element(xml, name);
  if (found === null) return null;
  return { size: numAttr(found, "w:w") ?? 0, type: attr(found, "w:type") };
}

/** The four sides of a `w:tcMar` or `w:tblCellMar`, in twips. */
function marginsOf(xml, name) {
  const found = element(xml, name);
  if (found === null) return null;

  return {
    top: childNum(found, "w:top", "w:w"),
    right: childNum(found, "w:right", "w:w"),
    bottom: childNum(found, "w:bottom", "w:w"),
    left: childNum(found, "w:left", "w:w"),
  };
}

/** Every edge of a `w:tblBorders` or `w:tcBorders`. An undrawn edge is null. */
function bordersOf(xml, name) {
  const found = element(xml, name);
  if (found === null) return null;

  return {
    top: borderOf(found, "top"),
    right: borderOf(found, "right"),
    bottom: borderOf(found, "bottom"),
    left: borderOf(found, "left"),
    insideH: borderOf(found, "insideH"),
    insideV: borderOf(found, "insideV"),
  };
}

/**
 * One cell, read whole.
 *
 * `column` is the grid column the cell begins in rather than its place among
 * its siblings. The two are the same number until something spans, and a case
 * about spanning that counted siblings would be counting the wrong thing.
 */
function readCell(xml, index, column) {
  const inner = innerOf(xml);
  const tcPr = element(headOf(inner, "w:p", "w:tbl"), "w:tcPr");
  const vMerge = tcPr === null ? null : element(tcPr, "w:vMerge");
  const gridSpan = childNum(tcPr, "w:gridSpan", "w:val");

  return {
    index,
    /** Where the cell begins in the table's grid, spans counted. */
    column,
    text: textOf(xml),

    /** The cell's own declared width, when it has one. */
    width: widthOf(tcPr, "w:tcW"),
    /** How many grid columns it runs across. */
    gridSpan: gridSpan ?? 1,
    /**
     * A vertical merge, in the two halves OOXML writes it in: `restart` opens
     * one, `continue` is a cell swallowed by the one above it. Null is a cell
     * that merges with nothing.
     */
    vMerge: vMerge === null ? null : attr(vMerge, "w:val") ?? "continue",

    /** Cell shading, as the fill hex without a `#`. */
    shd: childAttr(tcPr, "w:shd", "w:fill"),
    /** How the content sits against the height of the cell. */
    vAlign: childAttr(tcPr, "w:vAlign", "w:val"),
    /** Borders, by edge. An undrawn edge is null. */
    borders: bordersOf(tcPr, "w:tcBorders"),
    /** The room left inside the cell, in twips. */
    margins: marginsOf(tcPr, "w:tcMar"),

    /** The paragraphs the cell holds, read exactly as a body paragraph is. */
    paragraphs: scanElements(inner, "w:p", ["w:tbl"]).map(readParagraph),
    /** The tables inside it. A table that holds a table holds it here. */
    tables: scanElements(inner, "w:tbl").map((table, at) => readTable(table, at)),

    xml,
  };
}

/** One row, and the cells across it. */
function readRow(xml, index) {
  const inner = innerOf(xml);
  const trPr = element(headOf(inner, "w:tc"), "w:trPr");
  const height = trPr === null ? null : element(trPr, "w:trHeight");

  let column = 0;
  const cells = scanElements(inner, "w:tc", ["w:tbl"]).map((cell, at) => {
    const record = readCell(cell, at, column);
    column += record.gridSpan;
    return record;
  });

  return {
    index,
    text: textOf(xml),

    /**
     * Whether the row repeats at the top of every page the table runs onto.
     *
     * Three-state, like every OOXML toggle: `<w:tblHeader/>` is on,
     * `<w:tblHeader w:val="false"/>` is explicitly off, and no element at all
     * is inherit. The packer writes the middle one, and a probe that collapsed
     * it into the last would report a row as having said nothing when it said
     * no.
     */
    tblHeader: toggle(trPr, "w:tblHeader"),
    /** Whether the row is forbidden to split across a page. */
    cantSplit: toggle(trPr, "w:cantSplit"),

    /** A declared row height, in twips. */
    heightTwips: height === null ? null : numAttr(height, "w:val"),
    /** `atLeast` grows to fit, `exact` clips, `auto` takes the content's. */
    heightRule: height === null ? null : attr(height, "w:hRule") ?? "atLeast",

    cellCount: cells.length,
    cells,
  };
}

/** One table, and everything under it. */
function readTable(xml, index) {
  const inner = innerOf(xml);
  const head = headOf(inner, "w:tr");
  const tblPr = element(head, "w:tblPr");
  const grid = element(head, "w:tblGrid");
  const rows = scanElements(inner, "w:tr", ["w:tbl"]).map(readRow);

  return {
    index,
    text: textOf(xml),

    /** What the table declares its own width to be. */
    width: widthOf(tblPr, "w:tblW"),
    /**
     * How far the table is indented from the text column's left edge.
     *
     * Negative is a bleed — a table pulled out past the margin to the paper's
     * edge. `w:tblInd` carries its value in `w:w`, not the `w:left` a reader
     * of `w:ind` would look for, which is the whole reason the preview needed
     * a fix here.
     */
    indent: widthOf(tblPr, "w:tblInd"),
    /** Table alignment, when the table is not simply left where it stands. */
    jc: childAttr(tblPr, "w:jc", "w:val"),
    /** `fixed` honours the declared widths; `autofit` lets Word rework them. */
    layout: childAttr(tblPr, "w:tblLayout", "w:type"),
    /** A named Word table style, when the file leans on one. */
    style: childAttr(tblPr, "w:tblStyle", "w:val"),
    /** Whether the table floats, with the text wrapped around it. */
    floating: element(tblPr, "w:tblpPr") !== null,
    /** The grid every row's cells are laid on, in twips. */
    gridTwips: grid === null
      ? []
      : elements(grid, "w:gridCol").map((column) => numAttr(column, "w:w") ?? 0),
    /** The borders the table draws around and between its cells. */
    borders: bordersOf(tblPr, "w:tblBorders"),
    /** The room every cell leaves unless it says otherwise, in twips. */
    cellMargins: marginsOf(tblPr, "w:tblCellMar"),

    rowCount: rows.length,
    rows,
  };
}

/**
 * Every table in a list, the nested ones included, innermost first.
 *
 * `path` names where each one was found — `0` is the first body table, and
 * `0.1.2.0` is the first table inside the third cell of its second row. A
 * nested-table case is asking about exactly that string.
 *
 * The order is what makes an anchor mean the cell a reader would point at. A
 * cell's text is everything printed inside it, a nested table's words
 * included, so an outer cell matches every anchor its inner table matches —
 * and the outer one is never the answer, because the inner cell is the more
 * specific of the two. Deepest first settles it once, here, rather than in
 * every lookup.
 */
function flattenTables(tables, prefix = "", depth = 0) {
  const found = [];

  for (const table of tables) {
    const path = `${prefix}${table.index}`;

    for (const row of table.rows) {
      for (const cell of row.cells) {
        found.push(...flattenTables(cell.tables, `${path}.${row.index}.${cell.index}.`, depth + 1));
      }
    }

    found.push({ ...table, path, depth });
  }

  return found;
}

/**
 * Everything probe A reports about one packed file.
 *
 * @param {Uint8Array} zip The `.docx` bytes.
 * @returns The measurement, ready to be written as `measure-a.json`.
 */
export function measureOoxml(zip) {
  const document = partText(zip, "word/document.xml");
  const parts = partNames(zip);
  const styles = parts.includes("word/styles.xml") ? partText(zip, "word/styles.xml") : "";
  // Some things a document declares are settings rather than section
  // properties — `w:evenAndOddHeaders` among them — and live in their own part.
  // A case that looked for one in `document.xml` would find nothing and
  // conclude the document never asked for it.
  const settings = parts.includes("word/settings.xml") ? partText(zip, "word/settings.xml") : "";

  // Only body paragraphs: a paragraph inside a table cell belongs to the table
  // slice, and a header's belongs to the furniture slice. Both would otherwise
  // shift every index a paragraph case counts on.
  const body = element(document, "w:body") ?? document;
  const paragraphs = topLevelParagraphs(body).map(readParagraph);
  // The body's own tables, each read whole — its rows, its cells, and any
  // table one of those cells holds. The body slice above deliberately leaves a
  // cell's paragraphs out; this is where they are.
  const tables = scanElements(body, "w:tbl").map(readTable);


  const sectPr = element(body, "w:sectPr");
  const pgSz = sectPr === null ? null : element(sectPr, "w:pgSz");
  const pgMar = sectPr === null ? null : element(sectPr, "w:pgMar");

  return {
    probe: "A",
    parts,
    paragraphCount: paragraphs.length,
    paragraphs,
    /** Every paragraph in the file, cells and furniture included. */
    allParagraphCount: elements(document, "w:p").length,
    /** The body's tables, in order, each holding its own rows and cells. */
    tableCount: tables.length,
    tables,
    section: {
      widthTwips: pgSz === null ? null : numAttr(pgSz, "w:w"),
      heightTwips: pgSz === null ? null : numAttr(pgSz, "w:h"),
      orientation: pgSz === null ? null : attr(pgSz, "w:orient") ?? "portrait",
      margins: pgMar === null ? null : {
        top: numAttr(pgMar, "w:top"),
        right: numAttr(pgMar, "w:right"),
        bottom: numAttr(pgMar, "w:bottom"),
        left: numAttr(pgMar, "w:left"),
      },
      titlePg: sectPr === null ? null : toggle(sectPr, "w:titlePg"),
      // How far the running strips stand from the paper, in twips. A
      // different distance from the margin, and the one a document has no
      // way to state today.
      headerTwips: pgMar === null ? null : numAttr(pgMar, "w:header"),
      footerTwips: pgMar === null ? null : numAttr(pgMar, "w:footer"),
    },
    /** The default paragraph properties every paragraph inherits. */
    docDefaults: readDocDefaults(styles),
    documentXml: document,
    stylesXml: styles,
    settingsXml: settings,
  };
}

/**
 * The paragraphs of the body itself, with the ones inside tables left out.
 *
 * `<w:p>` nests: a cell holds paragraphs, and a naive scan returns a cell's
 * paragraph as if it stood in the body. Walking the body's immediate children
 * is what keeps a paragraph case's index meaning what it looks like it means.
 */
function topLevelParagraphs(body) {
  const found = [];
  let depth = 0;
  let start = -1;

  const pattern = /<w:(p|tbl)(?:\s[^>]*)?(\/?)>|<\/w:(p|tbl)>/g;

  for (const match of body.matchAll(pattern)) {
    const [whole, open, selfClose, close] = match;

    if (open === "tbl" || close === "tbl") {
      depth += close === "tbl" ? -1 : selfClose === "/" ? 0 : 1;
      continue;
    }

    if (depth > 0) continue;

    if (open === "p") {
      if (selfClose === "/") {
        // An empty paragraph, written closed. It still holds a line in Word.
        found.push(whole);
        continue;
      }
      if (start === -1) start = match.index;
    } else if (close === "p" && start !== -1) {
      found.push(body.slice(start, match.index + whole.length));
      start = -1;
    }
  }

  return found;
}

/** What `w:docDefaults` sets, which is what an unstyled paragraph inherits. */
function readDocDefaults(styles) {
  const defaults = element(styles, "w:docDefaults");
  if (defaults === null) return null;

  const pPrDefault = element(defaults, "w:pPrDefault");
  const rPrDefault = element(defaults, "w:rPrDefault");
  const spacing = pPrDefault === null ? null : element(pPrDefault, "w:spacing");
  const rPr = rPrDefault === null ? null : element(rPrDefault, "w:rPr");

  return {
    jc: pPrDefault === null ? null : childAttr(pPrDefault, "w:jc", "w:val"),
    spacing: spacing === null ? null : {
      before: numAttr(spacing, "w:before"),
      after: numAttr(spacing, "w:after"),
      line: numAttr(spacing, "w:line"),
      lineRule: attr(spacing, "w:lineRule"),
    },
    font: rPr === null ? null : childAttr(rPr, "w:rFonts", "w:ascii"),
    szHalfPt: rPr === null ? null : childNum(rPr, "w:sz", "w:val"),
    color: rPr === null ? null : childAttr(rPr, "w:color", "w:val"),
  };
}

/**
 * The view a case is handed: the measurement, plus the lookups that keep an
 * assertion to one line.
 */
export function ooxmlView(measure) {
  return {
    ...measure,

    /**
     * The paragraph whose text contains this anchor. First match wins.
     *
     * Case-insensitive: `w:caps` prints a label in capitals while leaving the
     * text alone, so the file and Word spell the same label two ways and one
     * anchor has to find both.
     */
    para(anchor) {
      const wanted = anchor.toLowerCase();
      return measure.paragraphs.find((p) => p.text.toLowerCase().includes(wanted)) ??
        missingParagraph(anchor);
    },

    /** The nth body paragraph, counting from zero. */
    paraAt(index) {
      return measure.paragraphs[index] ?? missingParagraph(`#${index}`);
    },

    /** Every paragraph whose text contains this anchor. */
    paras(anchor) {
      return measure.paragraphs.filter((p) => p.text.includes(anchor));
    },
    /**
     * One of the body's tables, wrapped so a row and a cell are one step away.
     *
     * `a.table(0).row(1).cell(2)` is the shape every table assertion is
     * written in. The wrapping happens here rather than in the measurement so
     * `measure-a.json` stays plain data — a reader of the evidence should not
     * have to know which of its fields were functions.
     */
    table(index = 0) {
      return tableView(measure.tables?.[index] ?? missingTable(`#${index}`));
    },

    /** Every cell whose text contains this anchor, nested tables included. */
    cells(anchor) {
      const wanted = anchor.toLowerCase();

      return flattenTables(measure.tables ?? []).flatMap((table) =>
        table.rows.flatMap((row) =>
          row.cells
            .filter((cell) => cell.text.toLowerCase().includes(wanted))
            .map((cell) => ({ ...cell, table: table.path, row: row.index }))
        )
      );
    },

    /** Every table in the file, the nested ones included, each with its path. */
    allTables() {
      return flattenTables(measure.tables ?? []);
    },

    /**
     * The cell whose text contains this anchor, wherever in the document it is.
     *
     * Nested tables included, and case-insensitively — a heading cell set in
     * `w:caps` says one thing in the file and another on the page, and an
     * anchor names what a reader sees.
     */
    cell(anchor) {
      const wanted = anchor.toLowerCase();

      for (const table of flattenTables(measure.tables ?? [])) {
        for (const row of table.rows) {
          for (const cell of row.cells) {
            if (cell.text.toLowerCase().includes(wanted)) {
              return { ...cell, table: table.path, row: row.index };
            }
          }
        }
      }

      return missingCell(anchor);
    },

  };
}


/** A table with its rows one step away, and its cells two. */
function tableView(table) {
  return {
    ...table,
    /** The nth row, counting from zero. */
    row(index = 0) {
      const found = table.rows?.[index] ?? missingRow(`#${index}`);
      return {
        ...found,
        /** The nth cell of that row, counting from zero. */
        cell(at = 0) {
          return found.cells?.[at] ?? missingCell(`#${index}.${at}`);
        },
      };
    },
  };
}

/**
 * The stand-ins for a table, a row and a cell that are not there.
 *
 * Every field is null or empty rather than the object being undefined, so an
 * assertion against a table the packer never wrote reports the property it
 * wanted as `null` instead of dying on a property access. "Not found" is a
 * measurement, and it is exactly the measurement an unsupported-feature case
 * exists to record.
 */
function missingTable(anchor) {
  return {
    index: -1,
    missing: true,
    anchor,
    text: "",
    width: null,
    indent: null,
    jc: null,
    layout: null,
    style: null,
    floating: false,
    gridTwips: [],
    borders: null,
    cellMargins: null,
    rowCount: 0,
    rows: [],
  };
}

function missingRow(anchor) {
  return {
    index: -1,
    missing: true,
    anchor,
    text: "",
    tblHeader: null,
    cantSplit: null,
    heightTwips: null,
    heightRule: null,
    cellCount: 0,
    cells: [],
  };
}

function missingCell(anchor) {
  return {
    index: -1,
    missing: true,
    anchor,
    column: null,
    text: "",
    width: null,
    gridSpan: null,
    vMerge: null,
    shd: null,
    vAlign: null,
    borders: null,
    margins: null,
    paragraphs: [],
    tables: [],
    xml: "",
  };
}
/**
 * The stand-in for a paragraph that is not there.
 *
 * Every field is null rather than the object being undefined, so an assertion
 * on a missing paragraph reports the property it wanted as `null` instead of
 * dying on a property access. "Not found" is a measurement.
 */
function missingParagraph(anchor) {
  return {
    index: -1,
    missing: true,
    anchor,
    text: "",
    style: null,
    jc: null,
    ind: null,
    spacing: null,
    keepNext: null,
    keepLines: null,
    widowControl: null,
    pageBreakBefore: null,
    tabs: [],
    shd: null,
    pBdr: null,
    run: {
      font: null,
      szHalfPt: null,
      color: null,
      bold: null,
      italics: null,
      caps: null,
      spacing: null,
    },
    hasDrawing: false,
    hasField: false,
    hasBreak: false,
    xml: "",
  };
}
