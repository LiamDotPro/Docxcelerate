/**
 * Probe A of the invoice verification harness: the file itself.
 *
 * Unzips `.verify/invoice.docx` and reads the OOXML facts the objectives
 * assert on — structure, styling, geometry as written, before any renderer
 * has an opinion. This is ground truth for what was packed: probes B and C
 * measure what a reader made of it, this probe measures what is actually
 * in the package.
 *
 * The zip reader is a port of `tests/docx.ts` (central directory +
 * `inflateRawSync`) — the same ~60 lines the test suite already trusts.
 * The XML is scanned with strings and regexes rather than a parser: the
 * producer is `docx` v9, whose output shape is stable and never splits
 * attributes across lines, and a dependency-free probe cannot be broken
 * by a dependency. Every fact a region cannot supply comes back `null`
 * (or an empty list) rather than a throw — the objectives treat `null`
 * as FAIL "not found", which is the contract's behaviour for a missing
 * region. See VERIFY-CONTRACT.md ("measure-a.json") for the schema.
 */
import { inflateRawSync } from "node:zlib";
import { readFile, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..", "..");
const VERIFY_DIR = resolve(ROOT, ".verify");

// ---------------------------------------------------------------------------
// Zip reading (ported from tests/docx.ts)
// ---------------------------------------------------------------------------

/** Reads one entry out of a zip, via its central directory. */
export function entryOf(zip, wanted) {
  const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  const text = new TextDecoder("latin1").decode(zip);

  // The end-of-central-directory record says where the directory starts. It is
  // last in the file, so it is found by scanning back for its signature.
  const end = text.lastIndexOf("PK");

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
      // The local header repeats the name and extra fields, at its own lengths.
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

/** Every entry name the zip's central directory lists, in directory order. */
export function partNamesOf(zip) {
  const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  const text = new TextDecoder("latin1").decode(zip);
  const end = text.lastIndexOf("PK");
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

// ---------------------------------------------------------------------------
// XML scanning primitives
// ---------------------------------------------------------------------------

/** The five entities the packer writes, decoded — `&amp;` last so it cannot double-decode. */
function decodeEntities(text) {
  return text
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

/** The text an XML fragment prints: every `w:t`, joined and decoded. */
function textOf(xml) {
  return decodeEntities(
    [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join(""),
  );
}

/** An attribute's value from a single tag's text, or null. */
function attrOf(tagXml, name) {
  if (!tagXml) return null;
  const match = tagXml.match(new RegExp(`${name}="([^"]*)"`));
  return match ? decodeEntities(match[1]) : null;
}

/** An attribute's value as a number, or null when absent. */
function numAttrOf(tagXml, name) {
  const value = attrOf(tagXml, name);
  return value === null ? null : Number(value);
}

/**
 * The first `<name .../>` or `<name ...>` tag in a fragment, as its raw text.
 * The name must be followed by a space, `/` or `>` so `w:p` cannot match
 * `w:pPr` and `w:tbl` cannot match `w:tblGrid`.
 */
function firstTagOf(xml, name) {
  if (!xml) return null;
  const match = xml.match(new RegExp(`<${name}(?: [^>]*)?/?>`));
  return match ? match[0] : null;
}

/** The first `<name>…</name>` span in a fragment (non-nesting elements only). */
function firstBlockOf(xml, name) {
  if (!xml) return null;
  const match = xml.match(new RegExp(`<${name}(?: [^>]*)?>[\\s\\S]*?</${name}>`));
  return match ? match[0] : null;
}

/**
 * Every top-level `<name>…</name>` element in a fragment, nesting-aware.
 * Needed because rows contain rows (through nested tables) and a lazy regex
 * would close the outer element at the inner element's close tag. Self-closing
 * occurrences count as complete elements.
 */
function elementsOf(xml, name) {
  const token = new RegExp(`<${name}(?: [^>]*)?>|</${name}>`, "g");
  const out = [];
  let depth = 0;
  let start = -1;

  for (const match of xml.matchAll(token)) {
    if (match[0].startsWith("</")) {
      depth -= 1;
      if (depth === 0 && start !== -1) {
        out.push(xml.slice(start, match.index + match[0].length));
        start = -1;
      }
    } else if (match[0].endsWith("/>")) {
      if (depth === 0) out.push(match[0]);
    } else {
      if (depth === 0) start = match.index;
      depth += 1;
    }
  }

  return out;
}

/**
 * Every `<w:tbl>` in a fragment — nested ones included, each as its own
 * range — with start/end offsets kept so a drawing's position can be tested
 * for containment (the "is this image in a card table" question).
 */
function tableRangesOf(xml) {
  const token = /<w:tbl(?: [^>]*)?>|<\/w:tbl>/g;
  const stack = [];
  const ranges = [];

  for (const match of xml.matchAll(token)) {
    if (match[0].startsWith("</")) {
      const start = stack.pop();
      if (start !== undefined) {
        ranges.push({ start, end: match.index + match[0].length });
      }
    } else {
      stack.push(match.index);
    }
  }

  return ranges
    .sort((a, b) => a.start - b.start)
    .map((range) => ({ ...range, xml: xml.slice(range.start, range.end) }));
}

/** A boolean OOXML toggle: present counts as on unless it says `w:val="false"`. */
function toggleOn(tagXml) {
  if (!tagXml) return false;
  const value = attrOf(tagXml, "w:val");
  return value !== "false" && value !== "0";
}

// ---------------------------------------------------------------------------
// Fact extraction
// ---------------------------------------------------------------------------

/** The run-level facts an objective asserts on: face, size, colour, tracking. */
function runFactsOf(runXml) {
  const rPr = firstBlockOf(runXml, "w:rPr");

  return {
    font: attrOf(firstTagOf(rPr, "w:rFonts"), "w:ascii"),
    sz: numAttrOf(firstTagOf(rPr, "w:sz"), "w:val"),
    color: attrOf(firstTagOf(rPr, "w:color"), "w:val"),
    caps: toggleOn(firstTagOf(rPr, "w:caps")),
    spacing: numAttrOf(firstTagOf(rPr, "w:spacing"), "w:val"),
    text: textOf(runXml),
  };
}

/** Per-side cell borders, `none` sides omitted so "no border" reads as `{}`. */
function bordersOf(tcPr) {
  const tcBorders = firstBlockOf(tcPr, "w:tcBorders");
  const borders = {};

  for (const side of ["top", "left", "bottom", "right"]) {
    const tag = firstTagOf(tcBorders, `w:${side}`);
    if (!tag) continue;
    const val = attrOf(tag, "w:val");
    if (val === "none" || val === "nil") continue;
    borders[side] = {
      val,
      color: attrOf(tag, "w:color"),
      sz: numAttrOf(tag, "w:sz"),
    };
  }

  return borders;
}

/** The facts of one table cell: fill, alignment, borders, runs, leading. */
function cellFactsOf(cellXml) {
  const tcPr = firstBlockOf(cellXml, "w:tcPr");
  const shdFill = attrOf(firstTagOf(tcPr, "w:shd"), "w:fill");
  const paras = elementsOf(cellXml, "w:p");
  const firstPara = paras[0] ?? null;
  const pPr = firstPara ? firstBlockOf(firstPara, "w:pPr") : null;
  const spacingTag = firstTagOf(pPr, "w:spacing");

  // A charge row's second paragraph is the muted note under the description.
  // F7 wants it set tighter than the line above it, so its spacing is
  // recorded separately from the cell's first paragraph.
  const notePara = paras[1] ?? null;
  const noteSpacingTag = notePara ? firstTagOf(firstBlockOf(notePara, "w:pPr"), "w:spacing") : null;
  const noteSpacing = notePara === null ? null : {
    line: numAttrOf(noteSpacingTag, "w:line"),
    after: numAttrOf(noteSpacingTag, "w:after"),
  };

  return {
    shd: shdFill === "auto" ? null : shdFill,
    vAlign: attrOf(firstTagOf(tcPr, "w:vAlign"), "w:val"),
    borders: bordersOf(tcPr),
    // Runs across the whole cell, text-bearing ones only — drawings are
    // reported through `images`, and empty runs assert nothing.
    runs: [...cellXml.matchAll(/<w:r(?: [^>]*)?>[\s\S]*?<\/w:r>/g)]
      .map((m) => runFactsOf(m[0]))
      .filter((run) => run.text.length > 0),
    paraSpacing: {
      line: numAttrOf(spacingTag, "w:line"),
      after: numAttrOf(spacingTag, "w:after"),
    },
    noteSpacing,
    text: firstPara === null ? "" : textOf(cellXml),
  };
}

/**
 * The facts of one table. The anchor is the first text the table prints —
 * how the objectives name a table without depending on document order,
 * which the fixes this harness measures will reshuffle.
 */
function tableFactsOf(tableXml) {
  const tblPr = firstBlockOf(tableXml, "w:tblPr");
  const grid = firstBlockOf(tableXml, "w:tblGrid");
  const firstText = elementsOf(tableXml, "w:p")
    .map((p) => textOf(p).trim())
    .find((t) => t.length > 0);

  return {
    anchor: firstText ? firstText.slice(0, 40) : null,
    tblInd: numAttrOf(firstTagOf(tblPr, "w:tblInd"), "w:w"),
    widthsTwips: grid
      ? [...grid.matchAll(/<w:gridCol [^>]*>/g)].map((m) => numAttrOf(m[0], "w:w"))
      : [],
    rows: elementsOf(tableXml, "w:tr").map((rowXml) => ({
      header: toggleOn(firstTagOf(firstBlockOf(rowXml, "w:trPr"), "w:tblHeader")),
      cells: elementsOf(rowXml, "w:tc").map(cellFactsOf),
    })),
  };
}

/**
 * Every `w:drawing` in a part, with its declared extent and whether it sits
 * inside a single-cell bordered table — F13's "card" packing. The innermost
 * enclosing table is the one that decides: a drawing in a cell of the
 * letterhead's three-column table is not in a card.
 */
function imagesOf(partXml, context) {
  const tables = tableRangesOf(partXml);

  return [...partXml.matchAll(/<w:drawing>[\s\S]*?<\/w:drawing>/g)].map((match) => {
    const extent = firstTagOf(match[0], "wp:extent");
    const enclosing = tables
      .filter((t) => t.start < match.index && match.index < t.end)
      .sort((a, b) => b.start - a.start)[0];

    let inCardTable = false;
    if (enclosing) {
      const rows = elementsOf(enclosing.xml, "w:tr");
      const cells = rows.length === 1 ? elementsOf(rows[0], "w:tc") : [];
      if (cells.length === 1) {
        const cellBorders = bordersOf(firstBlockOf(cells[0], "w:tcPr"));
        const tblBorders = firstBlockOf(firstBlockOf(enclosing.xml, "w:tblPr"), "w:tblBorders");
        const tblHasBorder = tblBorders
          ? /w:val="(?!none|nil)[^"]+"/.test(tblBorders)
          : false;
        inCardTable = Object.keys(cellBorders).length > 0 || tblHasBorder;
      }
    }

    return {
      context,
      hasDrawing: true,
      extentEmu: {
        cx: numAttrOf(extent, "cx") ?? 0,
        cy: numAttrOf(extent, "cy") ?? 0,
      },
      inCardTable,
    };
  });
}

/**
 * The footer facts F4/F5 need: the paragraph holding the PAGE field, its
 * alignment and spacing, and whether any paragraph in the part inlines a
 * drawing with text (the one-line footer's defining property).
 *
 * `\bPAGE\b` and not a substring test, because NUMPAGES contains PAGE.
 */
function footerPartFactsOf(partXml) {
  const paragraphs = elementsOf(partXml, "w:p");
  const fieldPara = paragraphs.find((p) =>
    [...p.matchAll(/<w:instrText[^>]*>([^<]*)</g)].some((m) => /\bPAGE\b/.test(m[1])),
  );

  let pageFieldPara = null;
  if (fieldPara) {
    const pPr = firstBlockOf(fieldPara, "w:pPr");
    const instr = [...fieldPara.matchAll(/<w:instrText[^>]*>([^<]*)</g)]
      .map((m) => m[1])
      .join(" ");
    pageFieldPara = {
      jc: attrOf(firstTagOf(pPr, "w:jc"), "w:val"),
      spacingAfter: numAttrOf(firstTagOf(pPr, "w:spacing"), "w:after"),
      hasPageField: /\bPAGE\b/.test(instr),
      hasNumPagesField: /\bNUMPAGES\b/.test(instr),
      inlineDrawingWithText: paragraphs.some(
        (p) => p.includes("<w:drawing>") && textOf(p).trim().length > 0,
      ),
    };
  }

  // F4 wants the credit on one line: the mark and its words in a single
  // paragraph, not a picture stacked above a caption.
  const drawingWithText = paragraphs.filter(
    (p) => p.includes("<w:drawing>") && textOf(p).trim().length > 0,
  );

  return {
    text: textOf(partXml),
    tables: tableRangesOf(partXml).map((range) => tableFactsOf(range.xml)),
    pageFieldPara,
    drawingWithTextParas: drawingWithText.length,
    drawingWithTextSample: drawingWithText.length === 0
      ? null
      : textOf(drawingWithText[0]).trim(),
  };
}

/** The sectPr's furniture references, resolved through the document rels. */
function sectPrFactsOf(documentXml, relTargets) {
  const matches = documentXml.match(/<w:sectPr(?: [^>]*)?>[\s\S]*?<\/w:sectPr>/g);
  if (!matches) return null;
  // The body's own sectPr is the last one — earlier ones belong to sections.
  const sectPr = matches[matches.length - 1];

  const refsOf = (name) =>
    [...sectPr.matchAll(new RegExp(`<w:${name} [^>]*/>`, "g"))].map((m) => ({
      type: attrOf(m[0], "w:type"),
      part: relTargets[attrOf(m[0], "r:id")] ?? null,
    }));

  return {
    titlePg: /<w:titlePg\b/.test(sectPr),
    headerRefs: refsOf("headerReference"),
    footerRefs: refsOf("footerReference"),
  };
}

/** Heading1's run facts from styles.xml — F10's tracking objective reads these. */
function heading1FactsOf(stylesXml) {
  const style = stylesXml.match(
    /<w:style [^>]*w:styleId="Heading1"[^>]*>[\s\S]*?<\/w:style>/,
  );
  if (!style) return null;
  const rPr = firstBlockOf(style[0], "w:rPr");

  return {
    sz: numAttrOf(firstTagOf(rPr, "w:sz"), "w:val"),
    // Run w:spacing (letter tracking), never the pPr's before/after spacing —
    // both share a tag name and only the rPr one is the F10 fact.
    spacing: numAttrOf(firstTagOf(rPr, "w:spacing"), "w:val"),
    caps: toggleOn(firstTagOf(rPr, "w:caps")),
    color: attrOf(firstTagOf(rPr, "w:color"), "w:val"),
  };
}

/**
 * The ids of every paragraph style that carries a page break.
 *
 * A break can ride the paragraph's style instead of the paragraph — the one
 * form Word and docx-preview both turn the page on. The fact worth measuring
 * is therefore "this paragraph breaks the page", however it comes to say so,
 * so the style ids are collected first and the paragraph scan consults them.
 */
function breakStyleIdsOf(stylesXml) {
  const ids = new Set();
  if (!stylesXml) return ids;

  for (const [, id, body] of stylesXml.matchAll(
    /<w:style [^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g,
  )) {
    const pPr = firstBlockOf(body, "w:pPr");
    if (pPr !== null && toggleOn(firstTagOf(pPr, "w:pageBreakBefore"))) ids.add(id);
  }

  return ids;
}

/**
 * The card table a body drawing sits in: its border ink and its cell padding.
 *
 * `imagesOf` already decides *whether* a drawing is in a card; F13 also asks
 * what the card is made of, which means going back to the same table. The
 * innermost enclosing table wins, exactly as it does there.
 */
function cardTableOf(documentXml) {
  const tables = tableRangesOf(documentXml);

  // The card is whichever drawing sits in a single-cell table of its own — not
  // simply the first picture in the file, which is the letterhead's mark and
  // sits in a three-column table with the sender's name beside it.
  let enclosing = null;
  let rows = [];
  let cells = [];

  for (const drawing of documentXml.matchAll(/<w:drawing>/g)) {
    const candidate = tables
      .filter((t) => t.start < drawing.index && drawing.index < t.end)
      .sort((a, b) => b.start - a.start)[0];
    if (!candidate) continue;

    const candidateRows = elementsOf(candidate.xml, "w:tr");
    const candidateCells = candidateRows.length === 1 ? elementsOf(candidateRows[0], "w:tc") : [];
    if (candidateCells.length !== 1) continue;

    enclosing = candidate;
    rows = candidateRows;
    cells = candidateCells;
    break;
  }

  if (enclosing === null) return null;

  const tcPr = firstBlockOf(cells[0], "w:tcPr");
  const cellBorders = bordersOf(tcPr);
  const tblBorders = firstBlockOf(firstBlockOf(enclosing.xml, "w:tblPr"), "w:tblBorders");

  // The ink can be written on the cell or on the table; either draws the card,
  // so whichever is present is the card's colour.
  const borderColor =
    Object.values(cellBorders).map((b) => b?.color).find((c) => c != null) ??
    attrOf(firstTagOf(tblBorders, "w:top"), "w:color") ??
    null;

  // Padding: the cell's own margins, else the table's default for all cells.
  const tcMar = firstBlockOf(tcPr, "w:tcMar") ??
    firstBlockOf(firstBlockOf(enclosing.xml, "w:tblPr"), "w:tblCellMar");

  return {
    borderColor: borderColor === "auto" ? null : borderColor,
    tcMarTwips: numAttrOf(firstTagOf(tcMar, "w:top"), "w:w"),
    widthTwips: numAttrOf(firstTagOf(firstBlockOf(enclosing.xml, "w:tblPr"), "w:tblW"), "w:w"),
    heightTwips: numAttrOf(firstTagOf(rows[0], "w:trHeight"), "w:val"),
  };
}

/**
 * Every paragraph style id that resolves to `baseId`: the id itself, plus any
 * style based on it, however many links down the chain. Order-independent, so
 * a style listed before its parent still resolves.
 */
function styleIdsResolvingTo(stylesXml, baseId) {
  const resolved = new Set([baseId]);
  if (!stylesXml) return resolved;

  const parentOf = new Map();
  for (const [, id, body] of stylesXml.matchAll(
    /<w:style [^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g,
  )) {
    const parent = attrOf(firstTagOf(body, "w:basedOn"), "w:val");
    if (parent !== null) parentOf.set(id, parent);
  }

  for (let grew = true; grew;) {
    grew = false;
    for (const [id, parent] of parentOf) {
      if (!resolved.has(id) && resolved.has(parent)) {
        resolved.add(id);
        grew = true;
      }
    }
  }

  return resolved;
}

/** Whether one paragraph turns the page before itself, directly or by style. */
function breaksBefore(paragraphXml, breakStyleIds) {
  const pPr = firstBlockOf(paragraphXml, "w:pPr");
  if (pPr === null) return false;
  if (toggleOn(firstTagOf(pPr, "w:pageBreakBefore"))) return true;

  const styleId = attrOf(firstTagOf(pPr, "w:pStyle"), "w:val");

  return styleId !== null && breakStyleIds.has(styleId);
}

/**
 * The longest run of adjacent paragraphs that print nothing — F12's other
 * fact. A break faked with empty paragraphs shows up here and nowhere else.
 */
function longestEmptyRun(paragraphs) {
  let longest = 0;
  let run = 0;

  for (const p of paragraphs) {
    const empty = textOf(p).trim() === "" && !p.includes("<w:drawing>");
    run = empty ? run + 1 : 0;
    if (run > longest) longest = run;
  }

  return longest;
}

/** Counts nodes of kind "section" anywhere in the built model. */
function sectionCountOf(node) {
  if (Array.isArray(node)) {
    return node.reduce((sum, item) => sum + sectionCountOf(item), 0);
  }
  if (node && typeof node === "object") {
    let sum = node.kind === "section" ? 1 : 0;
    for (const value of Object.values(node)) sum += sectionCountOf(value);
    return sum;
  }
  return 0;
}

// ---------------------------------------------------------------------------
// The probe
// ---------------------------------------------------------------------------

/**
 * Reads the packed invoice and writes `.verify/measure-a.json`.
 *
 * @param {{ docxPath?: string, modelPath?: string, outPath?: string }} options
 *   Paths, defaulted to the standard build outputs so the runner and the CLI
 *   agree without configuration; variant builds pass their own.
 * @returns The measurement object, as written.
 */
export async function measure({
  docxPath = resolve(VERIFY_DIR, "invoice.docx"),
  modelPath = resolve(VERIFY_DIR, "model.json"),
  outPath = resolve(VERIFY_DIR, "measure-a.json"),
} = {}) {
  const zip = new Uint8Array(await readFile(docxPath));
  const partNames = partNamesOf(zip).filter((name) => !name.endsWith("/"));
  const decoder = new TextDecoder();
  const partText = (name) =>
    partNames.includes(name) ? decoder.decode(entryOf(zip, name)) : null;

  const documentXml = partText("word/document.xml") ?? "";

  // r:id -> target basename, for resolving header/footer references.
  const relsXml = partText("word/_rels/document.xml.rels") ?? "";
  const relTargets = {};
  for (const match of relsXml.matchAll(/<Relationship [^>]*\/>/g)) {
    const id = attrOf(match[0], "Id");
    const target = attrOf(match[0], "Target");
    if (id && target) relTargets[id] = target.replace(/^.*\//, "");
  }

  const headerNames = partNames.filter((n) => /^word\/header\d*\.xml$/.test(n));
  const footerNames = partNames.filter((n) => /^word\/footer\d*\.xml$/.test(n));

  const headerParts = {};
  for (const name of headerNames) {
    const xml = partText(name);
    headerParts[name.replace("word/", "")] = {
      text: textOf(xml),
      hasTable: /<w:tbl(?: [^>]*)?>/.test(xml),
    };
  }

  const footerParts = {};
  for (const name of footerNames) {
    footerParts[name.replace("word/", "")] = footerPartFactsOf(partText(name));
  }

  // Body paragraphs, once — several fact groups walk the same list.
  const bodyParagraphs = elementsOf(documentXml, "w:p");

  const stylesXml = partText("word/styles.xml");
  const breakStyleIds = breakStyleIdsOf(stylesXml);

  // A section heading that also has to turn the page names a style based on
  // Heading1 rather than Heading1 itself, so "is this a section heading" is
  // asked of the style's ancestry, not its name.
  const heading1StyleIds = styleIdsResolvingTo(stylesXml, "Heading1");
  const headings1 = bodyParagraphs
    .filter((p) => {
      const id = attrOf(firstTagOf(firstBlockOf(p, "w:pPr"), "w:pStyle"), "w:val");
      return id !== null && heading1StyleIds.has(id);
    })
    .map((p) => textOf(p).trim());

  const images = [
    ...imagesOf(documentXml, "body"),
    ...headerNames.flatMap((n) => imagesOf(partText(n), "header")),
    ...footerNames.flatMap((n) => imagesOf(partText(n), "footer")),
  ];

  const allRenderedParts = ["word/document.xml", ...headerNames, ...footerNames];
  const imagePlaceholderTexts = allRenderedParts.flatMap((name) =>
    [...(partText(name) ?? "").matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)]
      .map((m) => decodeEntities(m[1]))
      .filter((t) => /^\[image: /.test(t)),
  );

  const pageBreaks = {
    // F12's fact: a paragraph whose whole contribution is a page break —
    // it prints nothing, draws nothing, and still costs a line.
    loneBreakParas: bodyParagraphs.filter(
      (p) =>
        /<w:br [^>]*w:type="page"/.test(p) &&
        textOf(p).trim() === "" &&
        !p.includes("<w:drawing>"),
    ).length,
    // Paragraphs that turn the page before themselves, whether they say so
    // directly or through a style that does.
    pageBreakBefore: bodyParagraphs
      .filter((p) => breaksBefore(p, breakStyleIds))
      .map((p) => textOf(p).trim().slice(0, 20)),
    breakStyleIds: [...breakStyleIds],
    maxConsecutiveEmptyParas: longestEmptyRun(bodyParagraphs),
  };

  // Left/right indents for every paragraph that prints text, anchored by its
  // first 20 characters — F14 reads the summary's right indent out of this.
  const inds = bodyParagraphs
    .map((p) => ({ p, text: textOf(p).trim() }))
    .filter(({ text }) => text.length > 0)
    .map(({ p, text }) => {
      const ind = firstTagOf(firstBlockOf(p, "w:pPr"), "w:ind");
      return {
        anchor: text.slice(0, 20),
        left: numAttrOf(ind, "w:left") ?? numAttrOf(ind, "w:start"),
        right: numAttrOf(ind, "w:right") ?? numAttrOf(ind, "w:end"),
      };
    });

  // Rounded geometry can land in any part (body, header, footer), so count
  // across everything the package renders — F15 accepts either mechanism.
  const allXmlText = partNames
    .filter((n) => n.endsWith(".xml") || n.endsWith(".rels"))
    .map((n) => partText(n) ?? "")
    .join("");
  const roundedGeometry = {
    roundrects: (allXmlText.match(/<v:roundrect\b[^>]*arcsize=/g) ?? []).length,
    prstRoundRects: (allXmlText.match(/prst="roundRect"/g) ?? []).length,
  };


  let sectionsInModel = null;
  try {
    sectionsInModel = sectionCountOf(JSON.parse(await readFile(modelPath, "utf8")));
  } catch {
    // No model to read is a missing region, not a probe crash.
  }

  const measurement = {
    parts: partNames,
    sectPr: sectPrFactsOf(documentXml, relTargets),
    headings1,
    headerParts,
    footerParts,
    bodyTables: tableRangesOf(documentXml).map((range) => tableFactsOf(range.xml)),
    images,
    imagePlaceholderTexts,
    pageBreaks,
    cardTable: cardTableOf(documentXml),
    inds,
    stylesXml: { heading1: stylesXml ? heading1FactsOf(stylesXml) : null },
    roundedGeometry,
    sectionsInModel,
  };

  await writeFile(outPath, `${JSON.stringify(measurement, null, 2)}\n`, "utf8");

  return measurement;
}

// CLI: node scripts/lib/probe-ooxml.mjs [--docx=path] [--model=path] [--out=path]
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const options = {};
  for (const arg of process.argv.slice(2)) {
    if (arg.startsWith("--docx=")) options.docxPath = resolve(arg.slice(7));
    if (arg.startsWith("--model=")) options.modelPath = resolve(arg.slice(8));
    if (arg.startsWith("--out=")) options.outPath = resolve(arg.slice(6));
  }
  measure(options).then((measurement) => {
    console.log(
      `measure-a: ${measurement.parts.length} parts, ` +
        `${measurement.bodyTables.length} body tables, ` +
        `${measurement.headings1.length} Heading1 paragraphs`,
    );
  }, (error) => {
    console.error("probe-ooxml failed:", error);
    process.exitCode = 1;
  });
}
