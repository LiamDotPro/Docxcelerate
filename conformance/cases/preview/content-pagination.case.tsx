import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * Whether the preview breaks a page when the text runs off the bottom of one.
 *
 * docx-preview does not: it breaks where the *file* says to — an explicit
 * `w:br type="page"` or a `w:pageBreakBefore` — and nowhere else, because it
 * has no paginator. So a document Word printed on five pages was drawn as one
 * very long sheet, and everything derived from the page count was wrong with
 * it: which page a paragraph is on, whether a header repeats, what a footer's
 * "1 of 5" should say. It was the largest single divergence in the suite, and
 * it blocked every case whose claim was about more than one page.
 *
 * `paginateDocxPreview` closes it. It flows the body's blocks into boxes the
 * height of the page and starts a new sheet when one is full, carrying the
 * running furniture onto it — from numbers the file already carries, which is
 * why it is finishing the reading rather than inventing an appearance.
 *
 * It is not Word's own paginator, and the difference is worth stating: it
 * breaks *between* blocks, never through the middle of one, so a paragraph Word
 * would split across a sheet moves whole here instead. On documents made of
 * ordinary paragraphs the two agree exactly, which is what this case measures.
 * A document of very long paragraphs is where they would drift, and the case to
 * write when one turns up.
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
  word: "Word repaginates continuously; docx-preview breaks only where the file says, so the framework paginates it",
  claim: "supported",

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
