import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A rule round every page, set in from the paper's edge.
 *
 * Until this landed the only frame was a bordered table, which frames the text
 * column and breaks across pages. Word's page border is measured from the
 * sheet and takes up no room; docx-preview parses `w:pgBorders` and draws
 * nothing, so the preview draws its own, and this measures where.
 */
export default defineCase({
  id: "page/border",
  feature: "page.border",
  title: "A double rule round the page, 10mm in from the paper's edge",
  word: "Design → Page Borders → Box, measured from edge of page (w:sectPr/w:pgBorders)",
  claim: "supported",

  style: {
    ...caseStyle,
    page: {
      ...caseStyle.page,
      border: "1F3A5F",
      borderStyle: "double",
      borderWidthPt: 1.5,
      borderOffsetMm: 10,
    },
  },

  document: template(
    <Document id="border" title="Page border">
      <Paragraph id="a">Certificate of completion</Paragraph>
    </Document>,
  ),

  regions: [{ id: "body", anchor: "Certificate" }],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.section.pgBordersOffsetFrom, "page", "measured from the paper");
      for (const side of ["top", "left", "bottom", "right"]) {
        const edge = a.section.pgBorders?.[side];
        is.equal(edge?.style, "double", `${side} is a double rule`);
        is.equal(edge?.color, "1F3A5F", `${side} colour`);
        is.equal(edge?.sizeEighthPt, 12, `${side} width`);
        is.equal(edge?.spacePt, 28, `${side} stands 28pt in`);
      }
    },
    preview: (b, is) => {
      const s = b.sections[0];
      is.equal(s.outline.style, "double", "the preview outlines the sheet");
      is.within(-s.outline.offset - s.outline.width, b.mm(10), "0.5mm", "the frame's outer edge stands 10mm in from the sheet");
    },
  },
});
