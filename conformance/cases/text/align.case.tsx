import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { COLUMN_MM, withBlocks } from "../_support/style.ts";

/**
 * A paragraph that is not set flush left.
 *
 * Until this landed it was impossible outside a table: `align` existed on a
 * column and on a cell, and `alignments` was reached only from `cellContent`,
 * so a body paragraph had no way to be centred, ranged right or justified. A
 * centred title over a letter, a right-ranged date, a justified block of terms
 * — none of them could be written, and the workaround was a single-celled
 * table, which is a table that exists to hold one alignment.
 *
 * Two ways in, because both are already how the framework works elsewhere: the
 * node says it (a date *is* ranged right, the way a money column is), and a
 * block can say it (a theme decides its `standfirst` is centred). The node
 * wins, exactly as a cell wins over its column.
 */
export default defineCase({
  id: "text/align",
  feature: "paragraph.align",
  title: "A paragraph set centred, ranged right, or justified",
  word: "Paragraph → Alignment (w:jc)",
  claim: "supported",

  style: withBlocks({
    /** A theme deciding what one of its named blocks looks like. */
    standfirst: { align: "center" },
  }),

  document: template(
    <Document id="align" title="Alignment">
      <Paragraph id="a">Left as written, which is what a paragraph does when it says nothing.</Paragraph>
      <Paragraph id="b" align="center">Centred by the node.</Paragraph>
      <Paragraph id="c" align="right">Ranged right by the node.</Paragraph>
      <Paragraph id="d" align="justify">
        Justified by the node, which needs enough words in it to show more than one
        line, because justification is a property of every line except the last and a
        single-line paragraph would look exactly like a left-ranged one.
      </Paragraph>
      <Paragraph id="e" variant="standfirst">Centred by the block.</Paragraph>
    </Document>
  ),

  regions: [
    { id: "left", anchor: "Left as written" },
    { id: "centre", anchor: "Centred by the node" },
    { id: "right", anchor: "Ranged right" },
    { id: "justify", anchor: "Justified by the node" },
    { id: "block", anchor: "Centred by the block" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.para("Left as written").jc, null, "an unaligned paragraph writes no w:jc at all");
      is.equal(a.para("Centred by the node").jc, "center", "centre writes w:jc center");
      is.equal(a.para("Ranged right").jc, "right", "right writes w:jc right");
      is.equal(a.para("Justified by the node").jc, "both", "justify writes w:jc both, which is what Word calls it");
      is.equal(a.para("Centred by the block").jc, "center", "a block's alignment reaches the paragraph too");
    },

    preview: (b, is) => {
      is.equal(b.para("Centred by the node").textAlign, "center", "the preview centres it");
      is.equal(b.para("Ranged right").textAlign, "right", "the preview ranges it right");
      is.equal(b.para("Justified by the node").textAlign, "justify", "the preview justifies it");

      // Geometry, not just the declared property: a centred line whose box is
      // full width is centred in name only, so the drawn text is measured.
      const column = b.mm(COLUMN_MM);
      is.within(b.textCentre("Centred by the node"), column / 2, "2mm", "the centred line sits on the column's axis");
      is.within(b.textRight("Ranged right"), column, "2mm", "the right-ranged line ends at the right margin");
      is.within(b.textLeft("Left as written"), 0, "1mm", "and the left one still starts at the left margin");
    },

    word: (c, is) => {
      is.equal(c.para("Centred by the node").alignment, "center", "Word reads it as centred");
      is.equal(c.para("Ranged right").alignment, "right", "Word reads it as right");
      is.equal(c.para("Justified by the node").alignment, "justify", "Word reads it as justified");
      is.equal(c.para("Left as written").alignment, "left", "and leaves the unaligned one alone");
    },

    parity: (p, is) => {
      is.within(p.previewX("Centred by the node"), p.wordX("Centred by the node"), "1mm", "the centred line starts where Word starts it");
      is.within(p.previewX("Ranged right"), p.wordX("Ranged right"), "1mm", "and so does the right-ranged one");
    },
  },
});
