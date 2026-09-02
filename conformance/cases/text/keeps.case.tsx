import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { withBlocks } from "../_support/style.ts";

/**
 * A paragraph that refuses to be split from what follows it, and one that
 * refuses to be split from itself.
 *
 * Neither is written today, so where a page break lands is entirely Word's
 * decision. A heading can be left alone at the foot of a page with its section
 * overleaf; a four-line address block can be cut after its first line. Both are
 * the kind of fault nobody sees in a preview of page one and everybody sees in
 * the printed copy — which is the argument for measuring it here.
 *
 * The document is deliberately long enough to reach a second page, and the
 * kept paragraphs are positioned so that a renderer ignoring the property
 * would split them. Asserting `w:keepNext` is present proves the file says it;
 * asking Word which page each landed on proves it *meant* something.
 */

/** Filler that pushes the kept pair towards the foot of page one. */
const filler = Array.from({ length: 26 }, (_, index) => (
  <Paragraph id={`f${index}`}>
    {`Body line ${index + 1}, present to consume the page so the kept pair meets the break.`}
  </Paragraph>
));

export default defineCase({
  id: "text/keeps",
  feature: "paragraph.keeps",
  title: "A paragraph kept with the next, and one kept whole",
  word: "Paragraph → Line and Page Breaks → Keep with next / Keep lines together",
  claim: "supported",

  style: withBlocks({
    /** A heading-like block: never the last thing on a page. */
    stayWithNext: { keepWithNext: true },
    /** A block that is one thing, and is not cut in half. */
    stayWhole: { keepLines: true },
  }),

  document: template(
    <Document id="keeps" title="Keeps">
      {filler}
      <Paragraph id="k" variant="stayWithNext">The heading that must not be orphaned.</Paragraph>
      <Paragraph id="k2">The paragraph it belongs to, which has to land on the same page.</Paragraph>
      <Paragraph id="w" variant="stayWhole">
        An address block, or anything else that is one thing rather than several: it runs to
        four or five lines, and a page break through the middle of it turns one object into
        two halves that each look like a mistake. Kept together, it moves to the next page
        whole, which is the only way it is ever correct.
      </Paragraph>
    </Document>
  ),

  regions: [
    { id: "keeper", anchor: "must not be orphaned" },
    { id: "kept", anchor: "has to land on the same page" },
    { id: "whole", anchor: "An address block" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.para("must not be orphaned").keepNext, true, "the block writes w:keepNext");
      is.equal(a.para("An address block").keepLines, true, "and the other writes w:keepLines");
      is.equal(a.para("Body line 1").keepNext, null, "an ordinary paragraph writes neither");
      is.equal(a.para("Body line 1").keepLines, null, "neither of them");
    },

    word: (c, is) => {
      is.equal(c.para("must not be orphaned").keepWithNext, true, "Word reads keep-with-next");
      is.equal(c.para("An address block").keepTogether, true, "Word reads keep-lines-together");

      // The point of the property, not the property itself.
      is.equal(
        c.para("must not be orphaned").page,
        c.para("has to land on the same page").page,
        "the heading and its paragraph land on one page",
      );
      is.equal(c.para("An address block").linesSplitAcrossPages, false, "and the block is not cut in half");
    },

    preview: (b, is) => {
      is.equal(
        b.para("must not be orphaned").pageIndex,
        b.para("has to land on the same page").pageIndex,
        "the preview keeps the pair together too",
      );
    },

    /**
     * The question this case is really about, now that it can be asked.
     *
     * Until the preview paginated, "both engines put the kept block on the same
     * page" could not pass and would have meant nothing if it had: docx-preview
     * drew the whole document as one sheet and reported page one for
     * everything. With pages, the keep is checkable on both sides — which is
     * the point, because a keep is a statement about where a page break may
     * fall and there is nothing to check without breaks.
     */
    parity: (p, is) => {
      is.equal(p.previewPages(), p.wordPages(), "both engines break the document into the same pages");
      is.equal(
        p.previewPage("must not be orphaned"),
        p.previewPage("has to land on the same page"),
        "the preview keeps the heading with its paragraph",
      );
      is.equal(
        p.previewPage("An address block"),
        p.wordPage("An address block"),
        "and puts the kept block on the page Word puts it on",
      );
    },
  },
});
