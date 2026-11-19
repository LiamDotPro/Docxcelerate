import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A list of what is covered and what is not, marked with ticks and crosses.
 *
 * `listMarker` is per item, so a tick and a cross stand in one list. Each
 * marker is a numbering of its own in the file, and the glyph is a plain
 * character in the paragraph's face: ✓ and ✗ are not in every body face, and
 * Word and Chrome both substitute one, which is the same choice the bullet
 * makes in not asking for a symbol font. The hang is held to the second tick
 * so a glyph the face lacks cannot move the words.
 */
export default defineCase({
  id: "lists/bullet-marker",
  feature: "list.bulletMarker",
  title: "A bulleted list marked with ticks, a cross and a dash",
  word: "Home → Bullets → Define New Bullet (w:lvlText)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="bullet-marker" title="Bullet marker">
      <Paragraph id="intro">What the policy covers:</Paragraph>
      <Paragraph id="fire" list="bullet" listMarker="tick">Fire and flood</Paragraph>
      <Paragraph id="theft" list="bullet" listMarker="tick">Theft from the home</Paragraph>
      <Paragraph id="wear" list="bullet" listMarker="cross">Wear and tear</Paragraph>
      <Paragraph id="note" list="bullet" listMarker="dash" listLevel={1}>Unless caused by an insured event</Paragraph>
      <Paragraph id="after">Ask us about anything else.</Paragraph>
    </Document>
  ),

  regions: [
    { id: "intro", anchor: "What the policy covers" },
    { id: "fire", anchor: "Fire and flood" },
    { id: "theft", anchor: "Theft from the home" },
    { id: "wear", anchor: "Wear and tear" },
    { id: "note", anchor: "Unless caused" },
    { id: "after", anchor: "Ask us about" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.list("Fire and flood").format, "bullet", "a marked item stands at a bullet level");
      is.equal(a.list("Fire and flood").text, "\u2713", "whose glyph is a tick");
      is.equal(a.list("Theft from the home").numId, a.list("Fire and flood").numId, "two ticks count in one numbering");
      is.equal(a.list("Wear and tear").text, "\u2717", "a cross has a numbering of its own");
      is.equal(a.list("Unless caused").text, "\u2013", "a dash is an en dash");
      is.equal(a.list("Unless caused").level, 1, "the dash stands one level down");
      is.greater(a.list("Fire and flood").hangingTwips, 0, "hung, so the glyph stands clear of the words");
    },

    preview: (b, is) => {
      is.includes(b.para("Fire and flood").marker, "\u2713", "the preview draws a tick in front of the item");
      is.includes(b.para("Wear and tear").marker, "\u2717", "and a cross in front of the other");
      is.within(
        b.para("Fire and flood").lines[0]?.x,
        b.para("Theft from the home").lines[0]?.x,
        "0.5mm",
        "and the words of both start in the same place",
      );
    },

    word: (c, is) => {
      // Word cannot read the glyph (see bulleted.case.tsx); the hang it can.
      is.greater(c.para("Fire and flood").leftIndent, 0, "Word sets the item in from the margin");
      is.less(c.para("Fire and flood").firstLineIndent, 0, "and hangs its first line, where the marker stands");
    },
  },
});
