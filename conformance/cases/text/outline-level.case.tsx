import { Document, Paragraph, Section, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { withBlocks } from "../_support/style.ts";

/**
 * Headings drawn from components, listed in Word's outline.
 *
 * Only a `<Section>` heading reached the navigation pane; a clause heading
 * drawn as a paragraph was body text to Word. An outline level draws no ink, so
 * there is no preview expectation: the claim is what the file says and what
 * Word reads back.
 */
export default defineCase({
  id: "text/outline-level",
  feature: "paragraph.outlineLevel",
  title: "A paragraph variant that is a heading in the outline",
  word: "Paragraph → Indents and Spacing → Outline level",
  claim: "supported",

  style: withBlocks({
    clauseHeading: { weight: "bold", outlineLevel: 1 },
    subClause: { outlineLevel: 2 },
  }),

  document: template(
    <Document id="outline-level" title="Outline level">
      <Paragraph id="c1" variant="clauseHeading">1. Definitions</Paragraph>
      <Paragraph id="c1a" variant="subClause">1.1 Interpretation</Paragraph>
      <Paragraph id="c1b">In this agreement the words have these meanings.</Paragraph>
      <Section id="annex" variant="subClause" title="Annex A">
        <Paragraph id="a1">The annex text.</Paragraph>
      </Section>
    </Document>
  ),

  regions: [
    { id: "clause", anchor: "1. Definitions" },
    { id: "annex", anchor: "Annex A" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.para("1. Definitions").outlineLvl, 0, "a level 1 block writes w:outlineLvl 0");
      is.equal(a.para("1.1 Interpretation").outlineLvl, 1, "level 2 writes 1");
      is.equal(a.para("In this agreement").outlineLvl, null, "body text writes none");
      is.equal(a.para("Annex A").style, "Heading1", "the section heading is still a Heading 1");
      is.equal(a.para("Annex A").outlineLvl, 1, "and sits one level down");
    },

    word: (c, is) => {
      is.equal(c.para("1. Definitions").outlineLevel, 1, "Word reads outline level 1");
      is.equal(c.para("1.1 Interpretation").outlineLevel, 2, "and level 2");
      is.equal(c.para("In this agreement").outlineLevel, 10, "body text is 10 to Word");
      is.equal(c.para("Annex A").outlineLevel, 2, "the annex heading is level 2");
    },
  },
});
