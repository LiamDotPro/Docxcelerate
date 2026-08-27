import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle, withBlocks } from "../_support/style.ts";

/**
 * The face, the size and the ink a paragraph is set in.
 *
 * The document says it once and a block may depart from it — a money column in
 * a face whose digits are all one width, a note set smaller and greyer than the
 * prose above it.
 *
 * The preview tier here carries a warning the others do not: a browser handed a
 * font name it cannot resolve falls back silently, and every measurement taken
 * against the wrong face is wrong by a plausible-looking amount. So this case
 * checks the face actually resolved before it believes anything else about the
 * page — see the `fontResolved` assertion.
 */
export default defineCase({
  id: "text/typography",
  feature: "paragraph.typography",
  title: "Face, size and ink, from the document and from a block",
  word: "Font → Font / Size / Font color (w:rFonts, w:sz, w:color)",
  claim: "supported",

  style: withBlocks({
    /** Figures that line up under one another. */
    money: { font: "Consolas", fontSizePt: 10 },
    /** A note set back from the prose it explains. */
    note: { color: "6B7280", fontSizePt: 8 },
    /** A line that carries emphasis on its own. */
    lead: { weight: "bold", fontSizePt: 14 },
  }),

  document: template(
    <Document id="typography" title="Typography">
      <Paragraph id="a">Body text, in the document's own face, size and ink.</Paragraph>
      <Paragraph id="b" variant="money">1,240.00</Paragraph>
      <Paragraph id="c" variant="note">A note, smaller and set back.</Paragraph>
      <Paragraph id="d" variant="lead">A line that leads.</Paragraph>
    </Document>
  ),

  regions: [
    { id: "body", anchor: "Body text, in the document's" },
    { id: "money", anchor: "1,240.00" },
    { id: "note", anchor: "A note, smaller" },
    { id: "lead", anchor: "A line that leads" },
  ],

  expect: {
    ooxml: (a, is) => {
      // The body's face and size live in w:docDefaults rather than on every
      // paragraph — writing them per run would be a file that says the same
      // thing a hundred times and disagrees with itself once.
      is.equal(a.docDefaults?.font, caseStyle.typography.bodyFont, "the document default names the body face");
      is.equal(a.docDefaults?.szHalfPt, caseStyle.typography.bodySizePt * 2, "and its size, in half-points");
      is.equal(a.docDefaults?.color, caseStyle.typography.color, "and its ink");

      is.equal(a.para("1,240.00").run.font, "Consolas", "the money block names its own face");
      is.equal(a.para("1,240.00").run.szHalfPt, 20, "and its own size");
      is.equal(a.para("A note, smaller").run.color, "6B7280", "the note block names its own ink");
      is.equal(a.para("A note, smaller").run.szHalfPt, 16, "and its own size");
      is.equal(a.para("A line that leads").run.bold, true, "the lead block is bold");
    },

    preview: (b, is) => {
      // First, and before anything else is believed: did the face resolve?
      // A measurement taken against a silent fallback is a measurement of the
      // wrong document.
      is.equal(b.fontResolved(caseStyle.typography.bodyFont), true, `${caseStyle.typography.bodyFont} resolves in the preview`);

      is.includes(b.para("Body text, in the document's").fontFamily, caseStyle.typography.bodyFont, "the body is set in the document's face");
      is.within(b.para("Body text, in the document's").fontSize, b.pt(caseStyle.typography.bodySizePt), "0.5px", "at the document's size");
      is.equal(b.para("Body text, in the document's").color, b.hex(caseStyle.typography.color), "in the document's ink");

      is.includes(b.para("1,240.00").fontFamily, "Consolas", "the money block draws in its own face");
      is.within(b.para("A note, smaller").fontSize, b.pt(8), "0.5px", "the note draws at its own size");
      is.equal(b.para("A note, smaller").color, b.hex("6B7280"), "and in its own ink");
      is.equal(b.para("A line that leads").fontWeight, "700", "the lead block draws bold");
    },

    word: (c, is) => {
      is.equal(c.para("Body text, in the document's").fontName, caseStyle.typography.bodyFont, "Word sets the body in the document's face");
      is.within(c.para("Body text, in the document's").fontSize, caseStyle.typography.bodySizePt, "0.1pt", "at its size");
      is.equal(c.para("1,240.00").fontName, "Consolas", "and the money block in its own");
      is.equal(c.para("A line that leads").bold, true, "and reads the lead block as bold");
    },
  },
});
