import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { withBlocks } from "../_support/style.ts";

/**
 * An epigraph in italic and a defined term in small capitals.
 *
 * Italic is `w:i` and small capitals are `w:smallCaps`; neither changes the
 * text, so anything reading the words back sees them as written. Word keeps
 * `w:caps` and `w:smallCaps` apart, so a small-capital run must carry no
 * `w:caps` for the two not to fight.
 */
export default defineCase({
  id: "text/italic-small-caps",
  feature: "paragraph.italicSmallCaps",
  title: "Italic and small capitals from a theme block",
  word: "Font → Italic + Small caps (w:i, w:smallCaps)",
  claim: "supported",

  style: withBlocks({
    epigraph: { italic: true },
    term: { transform: "smallCaps" },
  }),

  document: template(
    <Document id="italic-small-caps" title="Italic and small capitals">
      <Paragraph id="a" variant="epigraph">An epigraph</Paragraph>
      <Paragraph id="b" variant="term">Defined term</Paragraph>
    </Document>
  ),

  regions: [
    { id: "epigraph", anchor: "An epigraph" },
    { id: "term", anchor: "Defined term" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.para("An epigraph").run.italics, true, "the epigraph is italic by w:i");
      is.equal(a.para("Defined term").run.smallCaps, true, "the term is in small capitals by w:smallCaps");
      is.equal(a.para("Defined term").run.caps, null, "and carries no w:caps");
      is.equal(a.para("Defined term").text, "Defined term", "the text is still as written");
    },

    preview: (b, is) => {
      is.equal(b.para("An epigraph").fontStyle, "italic", "the preview sets the epigraph in italic");
      is.equal(b.para("Defined term").fontVariantCaps, "small-caps", "the preview draws small capitals");
      is.equal(b.para("Defined term").textTransform, "none", "without transforming the text");
    },

    word: (c, is) => {
      is.equal(c.para("An epigraph").italic, true, "Word prints the epigraph in italic");
    },
  },
});
