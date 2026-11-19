import { Document, PageNumber, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle, withBlocks } from "../_support/style.ts";

/**
 * Page numbers that start at iii and run in roman.
 *
 * The same document as `page-numbers`, with the page style asking for a
 * different start and a different way of writing it. Word's PAGE field follows
 * `w:pgNumType`, so the footer needs no change; the preview has to count from
 * the same place, and again after pagination, which is what this case checks.
 */

const lines = Array.from({ length: 90 }, (_, index) => (
  <Paragraph id={`p${index}`}>
    {`Body line ${index + 1} of ninety, enough to run this document well past one ` +
      "sheet so the numbering has something to count."}
  </Paragraph>
));

export default defineCase({
  id: "furniture/page-number-style",
  feature: "furniture.pageNumberStyle",
  title: "Page numbers that start elsewhere and run in roman",
  word: "Insert → Page Number → Format Page Numbers (w:pgNumType w:start, w:fmt)",

  claim: "supported",

  style: {
    ...withBlocks({ strip: { spacingAfterPt: 0 } }),
    page: { ...caseStyle.page, firstPageNumber: 3, numberStyle: "lowerRoman" },
  },

  document: template(
    <Document
      id="page-number-style"
      title="Page number style"
      footer={<PageNumber id="foot" variant="strip" format="currentOfTotal" separator=" / " />}
    >
      {lines}
    </Document>
  ),

  regions: [{ id: "foot", anchor: "/" }],

  expect: {
    ooxml: (a, is) => {
      is.includes(a.documentXml, '<w:pgNumType w:start="3" w:fmt="lowerRoman"/>', "the section numbers from iii");
    },

    word: (c, is) => {
      is.greater(c.pageCount(), 1, "the document runs to several pages");
      is.equal(
        /^\s*[ivxlcdm]+\s*\/\s*\d+\s*$/.test(c.furniture("primary", "footer").text ?? ""),
        true,
        "the footer reads a roman numeral over a count",
      );
    },

    preview: (b, is) => {
      is.includes(b.footerText(1), `iii / ${b.pageCount()}`, "page one says iii");
      is.includes(b.footerText(2), "iv / ", "and the refill after pagination says iv on page two");
    },

    parity: (p, is) => {
      is.includes(p.preview.footerText(1), `iii / ${p.wordPages()}`, "with the total Word will print");
    },
  },
});
