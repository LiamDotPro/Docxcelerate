import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * Whether the preview breaks a page when the text runs off the bottom of one.
 *
 * It does not. docx-preview breaks where the *file* says to break — an
 * explicit `w:br type="page"` or a `w:pageBreakBefore` — and nowhere else. It
 * does not lay text out against a page height and start a new sheet when the
 * old one is full, because that is a pagination engine and it does not have
 * one. So a document Word prints on three pages is drawn as one very long
 * sheet, and everything derived from the page count is wrong with it: which
 * page a paragraph is on, whether a kept block was moved whole, what the
 * footer's "1 / 3" should say.
 *
 * This is the largest single divergence between the two engines and it is not
 * a bug anyone can fix in a reading. Writing back a fact the file declares is
 * one thing; deciding where a line breaks across a sheet is a layout engine,
 * and building a second one is exactly what the framework refuses to do. So it
 * is written down, measured, and kept in front of us.
 *
 * The practical consequence, until this closes: **trust the preview for how a
 * page looks and Word for how many there are.** A document whose page breaks
 * matter should declare them rather than let them fall where they fall — which
 * is what `<PageBreak>` is for, and what the preview does honour.
 */

/** Enough prose to overrun an A4 page with 20mm margins several times over. */
const lines = Array.from({ length: 90 }, (_, index) => (
  <Paragraph id={`p${index}`}>
    {`Line ${index + 1} of ninety, long enough that ninety of them cannot fit on one A4 sheet ` +
      "however generously it is measured."}
  </Paragraph>
));

export default defineCase({
  id: "preview/content-pagination",
  feature: "preview.pagination",
  title: "The preview starts a new page when the text overruns one",
  word: "Word repaginates continuously; docx-preview breaks only where the file says",
  claim: "unsupported",

  style: caseStyle,

  document: template(
    <Document id="content-pagination" title="Content pagination">
      {lines}
    </Document>
  ),

  regions: [
    { id: "first", anchor: "Line 1 of ninety" },
    { id: "last", anchor: "Line 90 of ninety" },
  ],

  expect: {
    word: (c, is) => {
      // The control: Word does paginate, so there is something to disagree with.
      is.greater(c.pageCount(), 1, "Word breaks the text across several pages");
      is.greater(c.para("Line 90 of ninety").page, 1, "and the last line is not on page one");
    },

    preview: (b, is) => {
      is.greater(b.pageCount(), 1, "the preview breaks it too");
    },

    parity: (p, is) => {
      is.equal(p.previewPages(), p.wordPages(), "the two engines agree how many pages there are");
      is.equal(
        p.previewPage("Line 90 of ninety"),
        p.wordPage("Line 90 of ninety"),
        "and put the last line on the same one",
      );
    },
  },
});
