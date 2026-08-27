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

  // Only body paragraphs: a paragraph inside a table cell belongs to the table
  // slice, and a header's belongs to the furniture slice. Both would otherwise
  // shift every index a paragraph case counts on.
  const body = element(document, "w:body") ?? document;
  const paragraphs = topLevelParagraphs(body).map(readParagraph);

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
