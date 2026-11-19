import { Document, PageNumber, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { withBlocks } from "../_support/style.ts";

/**
 * "3 of 5" in the footer, counted by whatever lays the pages out.
 *
 * A build cannot know either number: how many pages a document runs to depends
 * on the paper, the face, and how much an engine wrote into every dynamic node.
 * So the document says which form it wants and the renderer counts — in Word
 * that is the PAGE and NUMPAGES fields, re-evaluated on every repagination.
 *
 * docx-preview drops a field run on the floor: its `renderRun` returns null for
 * one, so a footer Word prints as "Page 1 of 5" arrives as the words left
 * between the dropped fields. The words the document itself wrote — here the
 * separator — are the anchor, where the bare "/" used to be the only one.
 * `fillPageFields` puts the numbers back from
 * the pagination that just happened — which, until the preview paginated, meant
 * putting back "1 / 1" for a document Word printed on five sheets. This case
 * is the one that could not be honest before, and is the reason it is here now.
 */

const lines = Array.from({ length: 90 }, (_, index) => (
  <Paragraph id={`p${index}`}>
    {`Body line ${index + 1} of ninety, enough to run this document well past one ` +
      "sheet so the numbering has something to count."}
  </Paragraph>
));

export default defineCase({
  id: "furniture/page-numbers",
  feature: "furniture.pageNumbers",
  title: "A labelled page number the renderer counts, not the build",
  word: "Insert → Page Number (PAGE and NUMPAGES fields)",

  claim: "supported",

  style: withBlocks({ strip: { spacingAfterPt: 0 } }),

  document: template(
    <Document
      id="page-numbers"
      title="Page numbers"
      footer={<PageNumber id="foot" variant="strip" prefix="Page " separator=" of " align="right" />}
    >
      {lines}
    </Document>
  ),

  regions: [{ id: "foot", anchor: "Page" }],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.parts.includes("word/footer1.xml"), true, "the package holds a footer part");
    },

    word: (c, is) => {
      is.greater(c.pageCount(), 1, "the document runs to several pages");
      is.equal(c.furniture("primary", "footer").exists, true, "and carries a footer");
      // Word has evaluated the fields, so the strip reads as a number rather
      // than as an instruction.
      is.equal(
        /Page \d+ of \d+/.test(c.furniture("primary", "footer").text ?? ""),
        true,
        "whose text is a labelled number of a total",
      );
    },

    preview: (b, is) => {
      // Not the bare words docx-preview leaves behind, and not "1 of 1".
      is.equal(/^\s*Page\s+of\s*$/.test(b.footerText(1)), false, "the preview's footer is more than its words");
      is.includes(b.footerText(1), `Page 1 of ${b.pageCount()}`, "page one says one of however many there are");
      is.includes(
        b.footerText(b.pageCount()),
        `Page ${b.pageCount()} of ${b.pageCount()}`,
        "and the last page says the last of them",
      );
    },

    parity: (p, is) => {
      is.equal(p.previewPages(), p.wordPages(), "both engines count the same number of pages");
      is.includes(
        p.preview.footerText(1),
        `Page 1 of ${p.wordPages()}`,
        "so the preview's total is the total Word will print",
      );
    },
  },
});
