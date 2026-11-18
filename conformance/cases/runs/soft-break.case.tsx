import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A new line that is not a new paragraph.
 *
 * An address is one thing set on three lines. The only other break a document
 * can write is the one between paragraphs — which brings the paragraph's
 * space with it, so an address written as three paragraphs is three lines
 * standing as far apart as three separate thoughts, and three nodes
 * pretending to be one: three ids, three things for a condition to drop one
 * of.
 *
 * So the break is what an author reaches for first, a newline in the text,
 * and the packer writes it as the element that means one. It is the same
 * shape as the tab that `text/tab-stops` holds: a character left inside
 * `<w:t>` is whitespace, and the thing that says "break the line here" is
 * `<w:br/>`.
 *
 * **The preview tier was green while the file was wrong, and that is why the
 * file tier is the one to read here.** docx-preview sets every run
 * `white-space: pre-wrap`, so a newline left in `<w:t>` drew as a break on
 * screen with no break in the file behind it. The assertion that held this
 * case red was the count of `<w:br/>`. Whether Word broke the line before
 * there was one is unmeasured: this case was written on a machine without
 * Word, and its Word tier has still to run.
 *
 * The preview is asked about the lines' pitch rather than only their number.
 * Three lines is also what three paragraphs make, and the whole point of a
 * soft break is the space it does *not* bring — so the address's own lines
 * have to sit closer together than the address sits to the sentence under it.
 */
export default defineCase({
  id: "runs/soft-break",
  feature: "run.softBreak",
  title: "A new line that is not a new paragraph",
  word: "Shift+Enter — a manual line break (w:br inside a w:r)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="soft-break" title="Soft break">
      <Paragraph id="address">{"Flat 2\n14 Example Road\nLeeds LS1 4AB"}</Paragraph>
      <Paragraph id="after">A sentence under it, a paragraph's space away.</Paragraph>
    </Document>
  ),

  regions: [
    { id: "address", anchor: "Flat 2" },
    { id: "after", anchor: "A sentence under it" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.paragraphCount, 2, "the address is one paragraph, and the sentence under it another");
      is.includes(a.para("Flat 2").text, "Leeds LS1 4AB", "with all three of its lines in it");

      is.equal(a.para("Flat 2").xml.split("<w:br").length - 1, 2, "a break stands between each pair of lines");
      is.excludes(a.para("Flat 2").text, "\n", "and no newline is left in the text, where it is whitespace and not a break");
    },

    preview: (b, is) => {
      const address = b.para("Flat 2");
      const after = b.para("A sentence under it");

      is.equal(address.lineCount, 3, "the preview sets the address on three lines");
      is.less(
        address.lines[1]?.y - address.lines[0]?.y,
        after.lines[0]?.y - address.lines[2]?.y,
        "which sit closer to one another than a paragraph sits to the next",
      );
    },

    word: (c, is) => {
      is.equal(c.para("Flat 2").lineCount, 3, "Word sets the address on three lines");
      is.notEqual(c.para("A sentence under it").missing, true, "and the sentence under it is a paragraph of its own");
    },
  },
});
