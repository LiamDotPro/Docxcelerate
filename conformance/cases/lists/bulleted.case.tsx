import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A bulleted list, one item of which runs onto a second line and one of which
 * is a point under another.
 *
 * `list="bullet"` is what says it. A document that types its bullets into the
 * text gets paragraphs that start with a dot: a second line comes back to the
 * margin under the dot rather than under the words, and the items stand a
 * paragraph's space apart like four separate thoughts. This case was written
 * red, with the prop already on its paragraphs and nothing reading it.
 *
 * A list is three hops in the file, and each can be wrong alone. The
 * paragraph names a numbering; the numbering names an abstract one; the
 * abstract one holds the levels, where the glyph and the indent it hangs in
 * actually are. A paragraph with a `w:numPr` pointing at nothing is a
 * paragraph Word draws plain, so the assertions follow the chain to the level
 * rather than stopping at the element.
 *
 * The wrapped item is the one the geometry rests on. What makes a list a list
 * on the page is the hang: the words of every line start at the same place,
 * and the bullet stands to the left of all of them. A preview that drew the
 * glyph and then let the first line's words follow it directly would put the
 * first line a few points left of the second — so the two are held to each
 * other, which no count of bullets would catch.
 */
export default defineCase({
  id: "lists/bulleted",
  feature: "list.bullet",
  title: "A bulleted list, with an item that wraps and one nested under another",
  word: "Home → Bullets (w:numPr, w:numFmt bullet)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="bulleted" title="Bulleted list">
      <Paragraph id="intro">Bring these with you:</Paragraph>
      <Paragraph id="id" list="bullet">Photo identification</Paragraph>
      <Paragraph id="address" list="bullet">
        Proof of address dated within the last three months, such as a utility bill or a bank
        statement in your own name and showing the address you are applying from
      </Paragraph>
      <Paragraph id="council" list="bullet" listLevel={1}>A council tax bill also counts</Paragraph>
      <Paragraph id="after">We will return them the same day.</Paragraph>
    </Document>
  ),

  regions: [
    { id: "intro", anchor: "Bring these with you" },
    { id: "first", anchor: "Photo identification" },
    { id: "wrapped", anchor: "Proof of address" },
    { id: "nested", anchor: "A council tax bill" },
    { id: "after", anchor: "We will return them" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.paragraphCount, 5, "five paragraphs: a line, three items and a line");
      is.equal(a.para("Photo identification").text, "Photo identification", "an item's text is its words, with no bullet typed into it");

      // Followed the whole way to the level, not stopped at the element.
      is.equal(a.list("Photo identification").format, "bullet", "the first item stands at a bullet level");
      is.equal(a.list("Photo identification").text, "•", "whose glyph is a plain bullet, in no symbol font");
      is.greater(a.list("Photo identification").hangingTwips, 0, "hung, so the glyph stands clear of the words");
      is.equal(a.list("Proof of address").numId, a.list("Photo identification").numId, "the second item counts in the same list");
      is.equal(a.list("A council tax bill").level, 1, "the nested item stands one level down");
      is.greater(
        a.list("A council tax bill").leftTwips,
        a.list("Photo identification").leftTwips ?? Infinity,
        "and further in",
      );
    },

    preview: (b, is) => {
      const intro = b.para("Bring these with you");
      const first = b.para("Photo identification");
      const wrapped = b.para("Proof of address");
      const nested = b.para("A council tax bill");
      const after = b.para("We will return them");

      is.includes(first.marker, "•", "the preview draws a bullet in front of the item");
      is.equal(wrapped.lineCount, 2, "the long item runs onto a second line");
      is.greater(first.lines[0]?.x, intro.lines[0]?.x ?? Infinity, "an item's words start in from the margin");
      is.within(
        wrapped.lines[0]?.x,
        wrapped.lines[1]?.x,
        "0.5mm",
        "and its second line starts under the words, not under the bullet",
      );
      is.greater(nested.lines[0]?.x, first.lines[0]?.x ?? Infinity, "the nested item's words start further in again");
      is.less(
        nested.lines[0]?.y - wrapped.lines[1]?.y,
        after.lines[0]?.y - nested.lines[0]?.y,
        "and the items sit closer to one another than the list sits to what follows it",
      );
    },

    word: (c, is) => {
      // Word has no reading of the glyph here: the probe would need
      // `ListFormat.ListString`, and that is PowerShell nobody has run. What
      // it can already say is the hang, which is the half that shapes the page.
      is.greater(c.para("Photo identification").leftIndent, 0, "Word sets the item in from the margin");
      is.less(c.para("Photo identification").firstLineIndent, 0, "and hangs its first line, where the bullet stands");
      is.equal(c.para("Proof of address").lineCount, 2, "the long item runs onto a second line");
      is.greater(
        c.para("A council tax bill").leftIndent,
        c.para("Photo identification").leftIndent ?? Infinity,
        "and the nested item is set further in",
      );
    },
  },
});
