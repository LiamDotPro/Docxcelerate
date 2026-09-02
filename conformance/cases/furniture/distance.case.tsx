import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle, withBlocks } from "../_support/style.ts";

/**
 * How far the running strips stand from the edge of the paper.
 *
 * Until this landed a document could not say. The packer never wrote
 * `w:pgMar/@w:header` or `@w:footer`, so every document got the packing
 * library's own default of 708 twips — 12.5mm — whatever its margins were. That
 * number was chosen by a dependency, and it was the one piece of a document's
 * page geometry nobody here had decided.
 *
 * It matters more than it sounds. The header distance is measured from the
 * paper, not from the margin, so it is what decides whether a letterhead clears
 * a printer's unprintable edge, whether a running strip collides with a punched
 * hole, and how much air stands between the strip and the first line of text. A
 * document set with 25mm margins gets a header 12.5mm from the top; one set
 * with 10mm margins gets a header in exactly the same place, sitting almost on
 * the body.
 *
 * The style says everything else about the page in millimetres, so these do
 * too: `page.headerMm` and `page.footerMm`, beside the margins they are
 * measured against. A document that says nothing still gets 12.5mm, because
 * changing that would move the furniture of every document that never asked.
 */
export default defineCase({
  id: "furniture/distance",
  feature: "furniture.distance",
  title: "How far the running strips stand from the paper's edge",
  word: "Layout → Margins → Layout → From edge: Header / Footer (w:pgMar w:header, w:footer)",
  claim: "supported",

  style: {
    ...withBlocks({ strip: { spacingAfterPt: 0 } }),
    page: {
      ...caseStyle.page,
      // Measured from the paper, not the margin: 20mm margins with a
      // letterhead 15mm down, which is a different distance and now sayable.
      headerMm: 15,
      footerMm: 18,
    },
  },

  document: template(
    <Document
      id="distance"
      title="Furniture distance"
      header={<Paragraph id="h" variant="strip">A letterhead, fifteen millimetres down</Paragraph>}
      footer={<Paragraph id="f" variant="strip">A footer, eighteen millimetres up</Paragraph>}
    >
      <Paragraph id="a">Body.</Paragraph>
    </Document>
  ),

  regions: [
    { id: "head", anchor: "A letterhead" },
    { id: "foot", anchor: "A footer" },
  ],

  expect: {
    ooxml: (a, is) => {
      // 15mm is 850 twips, 18mm is 1020. What is written today is 708 for both,
      // which is neither, and is 12.5mm because that is what `docx` defaults to.
      is.within(a.section.headerTwips, 15 * 56.6929, 2, "the section declares a 15mm header distance");
      is.within(a.section.footerTwips, 18 * 56.6929, 2, "and an 18mm footer distance");
    },

    word: (c, is) => {
      is.within(c.headerDistance(), c.mm(15), "0.5mm", "Word puts the header 15mm from the top");
      is.within(c.footerDistance(), c.mm(18), "0.5mm", "and the footer 18mm from the foot");
    },

    preview: (b, is) => {
      is.within(b.furniture("header").y, b.mm(15), "1mm", "the preview draws the letterhead 15mm down");
    },

    parity: (p, is) => {
      is.within(p.previewHeaderY(), p.wordHeaderY(), "1mm", "and puts it where Word puts it");
    },
  },
});
