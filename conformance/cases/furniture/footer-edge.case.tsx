import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle, withBlocks } from "../_support/style.ts";

/**
 * A footer bar standing on the edge of the paper.
 *
 * `w:pgMar/@w:footer` is the distance from the bottom of the sheet to the
 * *bottom* of the footer, so a footer grows upward from that line. Set it to
 * zero and Word draws the strip flush with the paper's edge — which is what a
 * full-bleed bar is for, and what `furniture/distance` never checked, because
 * the footer there is taller than the room left under it and lands in the same
 * place whichever end it is measured from.
 *
 * This is the other case, and the one a real invoice hits: a strip *shorter*
 * than the room reserved for it. docx-preview gives a footer a box the height
 * of the margin it has to play with and lays the content out from the top of
 * it, so a short bar floats above where Word puts it — six pixels of white
 * under a strip that reaches every other edge of the paper. `settleDocxPreview`
 * seats it, and this is the assertion that says it is seated where Word seats
 * it rather than merely lower than it was.
 *
 * The margin is deliberately generous. A tight one leaves no reserve for a
 * short bar to float inside, which is exactly how this went unnoticed.
 */
export default defineCase({
  id: "furniture/footer-edge",
  feature: "furniture.footerEdge",
  title: "A footer bar standing on the paper's edge",
  word: "Layout → Margins → Layout → From edge: Footer 0 (w:pgMar w:footer)",
  claim: "supported",

  style: {
    ...withBlocks({
      /** A strip of colour, run to both edges of the sheet. */
      bar: {
        fill: "1E2A66",
        color: "FFFFFF",
        bleed: true,
        paddingPt: 10,
        fontSizePt: 8,
        spacingAfterPt: 0,
      },
    }),
    page: {
      ...caseStyle.page,
      // Room under the bar for it to float in, if anything is going to let it.
      margins: { topMm: 20, rightMm: 20, bottomMm: 25, leftMm: 20 },
      footerMm: 0,
    },
  },

  document: template(
    <Document
      id="footer-edge"
      title="Footer edge"
      footer={<Paragraph id="f" variant="bar">A bar on the edge of the paper</Paragraph>}
    >
      <Paragraph id="a">Body.</Paragraph>
    </Document>
  ),

  regions: [{ id: "foot", anchor: "A bar on the edge" }],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.section.footerTwips, 0, "the section asks for a footer on the paper's edge");
    },

    word: (c, is) => {
      is.within(c.footerDistance(), 0, "0.5mm", "Word stands the footer on the edge");
    },

    preview: (b, is) => {
      // The ink, not the box docx-preview reserved for it: the box was always
      // flush, and the bar inside it was not.
      is.within(b.furniture("footer").fromBottom, 0, "1mm", "the preview draws the bar on the edge too");
    },

    parity: (p, is) => {
      is.within(
        p.previewFooterFromBottom(),
        p.wordFooterFromBottom(),
        "1mm",
        "and puts it where Word puts it",
      );
    },
  },
});
