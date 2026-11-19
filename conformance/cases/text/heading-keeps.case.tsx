import { Document, Paragraph, Section, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A section heading that refuses to be left alone at the foot of a page.
 *
 * The keep is written on the Heading 1 style rather than on the paragraph, so
 * the OOXML assertions look at the style and the Word assertions ask Word what
 * it resolved. The filler is sized so that, were the keep ignored, the heading
 * would be the last block on page one with its paragraph overleaf.
 */

/** Filler that pushes the heading to the foot of page one. */
const filler = Array.from({ length: 27 }, (_, index) => (
  <Paragraph id={`f${index}`}>
    {`Body line ${index + 1}, present to consume the page so the heading meets the break.`}
  </Paragraph>
));

export default defineCase({
  id: "text/heading-keeps",
  feature: "section.heading.keeps",
  title: "A section heading kept with the paragraph after it",
  word: "Heading 1 → Line and Page Breaks → Keep with next / Keep lines together",
  claim: "supported",
  style: caseStyle,

  document: template(
    <Document id="heading-keeps" title="Heading keeps">
      {filler}
      <Section id="s" title="The heading that must not be stranded">
        <Paragraph id="p">The paragraph it introduces, which has to land on the same page.</Paragraph>
      </Section>
    </Document>
  ),

  regions: [
    { id: "heading", anchor: "must not be stranded" },
    { id: "body", anchor: "has to land on the same page" },
  ],

  // No `preview` or `parity` page assertion: the paginator does not consult
  // keeps, so one would only pass by accident. Word is the authority here.
  expect: {
    ooxml: (a, is) => {
      is.equal(a.para("must not be stranded").style, "Heading1", "the heading uses the Heading 1 style");
      is.equal(a.para("must not be stranded").keepNext, null, "and carries no keep itself: it rides the style");
      const heading = a.stylesXml.match(/<w:style [^>]*w:styleId="Heading1"[\s\S]*?<\/w:style>/)?.[0] ?? "";
      is.includes(heading, "<w:keepNext/>", "the style keeps with next");
      is.includes(heading, "<w:keepLines/>", "and keeps its lines together");
    },

    word: (c, is) => {
      is.equal(c.para("must not be stranded").keepWithNext, true, "Word resolves keep-with-next from the style");
      is.equal(c.para("must not be stranded").keepTogether, true, "and keep-lines-together");
      is.equal(
        c.para("must not be stranded").page,
        c.para("has to land on the same page").page,
        "the heading lands on its paragraph's page",
      );
    },
  },
});
