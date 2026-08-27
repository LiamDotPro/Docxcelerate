import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { withBlocks } from "../_support/style.ts";

/**
 * Left-hand pages and right-hand pages, with furniture of their own.
 *
 * A document that will be printed on both sides and bound has two kinds of
 * page, not one. The reference belongs at the outside edge of each — right on a
 * recto, left on a verso — so it is always the corner a thumb reaches, and a
 * folio that sat in the same place on every sheet would sit in the gutter on
 * half of them. It is the difference between a report and a printed report.
 *
 * Word calls it `w:evenAndOddHeaders`, and it is a document-level setting
 * rather than a section one: on, and every section takes its even pages from
 * the `even` parts. The framework has no way to ask. `header` and `footer` are
 * the running pair, `firstHeader` and `firstFooter` are page one's, and there
 * is no third pair — so `evenHeader` and `evenFooter` alongside them, which is
 * the shape the model already uses for the first page.
 *
 * Mirrored margins are the other half of the same job and a separate question:
 * this case is only about the strips.
 */

const lines = Array.from({ length: 60 }, (_, index) => (
  <Paragraph id={`p${index}`}>
    {`Body line ${index + 1}, here to carry the document onto a verso so the two ` +
      "kinds of page have something to differ about."}
  </Paragraph>
));

export default defineCase({
  id: "furniture/even-odd",
  feature: "furniture.evenOdd",
  title: "Left-hand and right-hand pages with furniture of their own",
  word: "Header & Footer → Different Odd & Even Pages (w:evenAndOddHeaders)",
  claim: "unsupported",

  style: withBlocks({ strip: { spacingAfterPt: 0 } }),

  document: template(
    <Document
      id="even-odd"
      title="Even and odd furniture"
      header={<Paragraph id="h" variant="strip">Fernhill Systems — recto</Paragraph>}
      evenHeader={<Paragraph id="eh" variant="strip">verso — Fernhill Systems</Paragraph>}
      footer={<Paragraph id="f" variant="strip">Registered in England</Paragraph>}
    >
      {lines}
    </Document>
  ),

  regions: [
    { id: "recto", anchor: "recto" },
    { id: "verso", anchor: "verso" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.includes(
        a.documentXml,
        "<w:evenAndOddHeaders",
        "the document turns on even and odd headers",
      );
      is.includes(a.documentXml, 'w:type="even"', "and the section points at an even-page header");

      const headers = a.parts.filter((part) => /^word\/header\d+\.xml$/.test(part));
      is.equal(headers.length, 2, "the package holds a header part for each kind of page");
    },

    word: (c, is) => {
      is.equal(c.differentOddAndEven(), true, "Word reads the document as having two kinds of page");
      is.equal(c.furniture("evenPages", "header").exists, true, "and a header for the even ones");
      is.includes(c.furniture("evenPages", "header").text, "verso", "carrying the verso strip");
      is.includes(c.furniture("primary", "header").text, "recto", "while the odd ones carry the recto");
    },

    preview: (b, is) => {
      is.includes(b.headerText(1), "recto", "page one is a recto");
      is.includes(b.headerText(2), "verso", "and page two is a verso");
    },
  },
});
