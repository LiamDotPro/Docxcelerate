import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle, withBlocks } from "../_support/style.ts";

/**
 * The body set mid-page by the page's own vertical alignment.
 *
 * Until this landed a one-page certificate or cover sheet always started at the
 * top margin, and the only way down was empty paragraphs that stop being right
 * the moment the text changes length. Word has a setting for it,
 * `w:sectPr/w:vAlign`; docx-preview skips it, so the preview has to put the
 * body where Word does, and this is what measures that they agree.
 */
export default defineCase({
  id: "page/vertical-align",
  feature: "page.valign",
  title: "The body centred between the top and bottom margins",
  word: "Layout → Page Setup → Layout → Vertical alignment: Center (w:sectPr/w:vAlign)",
  claim: "supported",

  // No space after, so the box measured is the line and nothing hangs under it.
  style: {
    ...withBlocks({ line: { spacingAfterPt: 0 } }),
    page: { ...caseStyle.page, valign: "center" },
  },

  document: template(
    <Document id="valign" title="Vertical alignment">
      <Paragraph id="a" variant="line">Certificate of completion</Paragraph>
    </Document>,
  ),

  regions: [{ id: "body", anchor: "Certificate" }],

  expect: {
    ooxml: (a, is) => is.equal(a.section.vAlign, "center", "the section declares w:vAlign center"),
    // A4 less 20mm margins top and bottom is a 257mm text area.
    preview: (b, is) => {
      const p = b.para("Certificate");
      is.within(p.y + p.h / 2, b.sections[0].content.h / 2, "1mm", "the preview draws it mid-page");
    },
    word: (c, is) =>
      is.within(c.para("Certificate").y, c.mm(257 / 2), "4mm", "Word sets it about half way down"),
    parity: (p, is) =>
      is.within(
        p.previewY("Certificate"),
        p.wordY("Certificate"),
        "1mm",
        "and the preview puts it where Word does",
      ),
  },
});
