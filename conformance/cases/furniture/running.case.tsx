import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle, withBlocks } from "../_support/style.ts";

/**
 * A strip along the top of every page, and one along the bottom.
 *
 * Running furniture is the part of a document nobody reads and everybody
 * notices when it is wrong: a reference at the head of page four, a company
 * number at the foot of every sheet. It is also the part a preview is least
 * likely to get right, because it sits *outside* the margins — a header is not
 * the first thing in the body, it is drawn in the space above it, and the two
 * are measured from different edges.
 *
 * Enough body to reach a second page, because "on every page" is the whole
 * claim and a one-page document cannot test it.
 */

const lines = Array.from({ length: 60 }, (_, index) => (
  <Paragraph id={`p${index}`}>
    {`Body line ${index + 1}, present to push this document onto a second page so ` +
      "the running strips have more than one page to run along."}
  </Paragraph>
));

export default defineCase({
  id: "furniture/running",
  feature: "furniture.running",
  title: "A header and a footer, drawn on every page",
  word: "Insert → Header / Footer (w:headerReference, w:footerReference)",
  claim: "supported",

  /**
   * The strips close up under themselves.
   *
   * The document's 10pt space-after applies inside a running strip like
   * anywhere else, and below the last line of a footer it is 10pt of nothing
   * between the words and the edge Word measures the footer from. Leaving it in
   * would make this case's geometry a question about paragraph spacing wearing
   * a footer's clothes; `text/spacing-after` already asks that one properly.
   */
  style: withBlocks({ strip: { spacingAfterPt: 0 } }),

  document: template(
    <Document
      id="running"
      title="Running furniture"
      header={
        <Paragraph id="head" variant="strip">INV-2026-0142 · Fernhill Systems Ltd</Paragraph>
      }
      footer={
        <Paragraph id="foot" variant="strip">Registered in England, number 09182736</Paragraph>
      }
    >
      {lines}
    </Document>
  ),

  regions: [
    { id: "head", anchor: "INV-2026-0142" },
    { id: "foot", anchor: "Registered in England" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.parts.includes("word/header1.xml"), true, "the package holds a header part");
      is.equal(a.parts.includes("word/footer1.xml"), true, "and a footer part");

      // Written once, referenced from the section. A strip repeated into the
      // body would be text that says the same words in the wrong place, and
      // would not repeat when the document grew a page.
      is.includes(a.documentXml, '<w:headerReference w:type="default"', "the section points at the header");
      is.includes(a.documentXml, '<w:footerReference w:type="default"', "and at the footer");

      is.equal(
        a.paras("INV-2026-0142").length,
        0,
        "the running text is not in the body — it belongs to the part",
      );
    },

    word: (c, is) => {
      is.greater(c.pageCount(), 1, "the document runs to more than one page");
      is.equal(c.furniture("primary", "header").exists, true, "Word has a header for every page");
      is.includes(c.furniture("primary", "header").text, "INV-2026-0142", "saying what the document said");
      is.equal(c.furniture("primary", "footer").exists, true, "and a footer");
      is.includes(c.furniture("primary", "footer").text, "Registered in England", "saying what the document said");

      // Not a title page: this document asked for the same strip everywhere.
      is.equal(c.differentFirstPage(), false, "and no first page of its own");
    },

    preview: (b, is) => {
      is.equal(b.furniture("header").drawn, true, "the preview draws the header");
      is.includes(b.headerText(), "INV-2026-0142", "with the document's words in it");
      is.equal(b.furniture("footer").drawn, true, "and the footer");
      is.includes(b.footerText(), "Registered in England", "with the document's words in it");

      // Outside the text column, which is the whole point of a running strip.
      is.less(b.furniture("header").y, b.mm(caseStyle.page.margins.topMm), "the header sits above the top margin");
    },

    /**
     * Where the strips sit against the paper.
     *
     * The header is a clean comparison: Word's `HeaderDistance` is where it
     * puts the top of the strip, and the top of the drawn text is what the
     * preview reports, and a header's first line starts at the top of its
     * region in both.
     *
     * The footer is not, and it is worth saying why rather than inventing a
     * number that passes. `FooterDistance` is where Word puts the *bottom of
     * the footer region*; what the preview can be asked for is the bottom of
     * the drawn line box. Those are different edges — measured, about 2mm apart
     * on one line of 11pt text — and the gap between them is the line box's own
     * descent, not a disagreement about the document. Comparing them would
     * report a divergence that is not there, and widening the tolerance until
     * it passed would hide a real one later.
     *
     * Closing it means measuring where Word actually draws the footer's glyphs,
     * which means selecting a range inside the footer story, which means moving
     * the window's `SeekView` — and that call kills a hidden Word instance
     * outright, taking the PDF export with it. So the footer is held to what
     * can be asked honestly: it is drawn, it says the right words, and it is
     * below the body.
     */
    parity: (p, is) => {
      is.within(p.previewHeaderY(), p.wordHeaderY(), "1mm", "the header sits where Word puts it");
      is.equal(p.previewPages(), p.wordPages(), "both engines break the document into the same pages");
      // The claim in the title, checked at last: the strips are on every sheet
      // the preview drew, not only the first. This could not be asked until
      // `preview/content-pagination` closed, because there was only ever one.
      is.equal(
        p.preview.furniture("header", p.previewPages()).drawn,
        true,
        "and the header is drawn on the last page as well as the first",
      );
      is.equal(
        p.preview.furniture("footer", p.previewPages()).drawn,
        true,
        "and so is the footer",
      );
    },
  },
});
