import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { withBlocks } from "../_support/style.ts";

/**
 * A first page whose furniture is its own.
 *
 * A letter's letterhead *is* the top of page one; the running strip that names
 * the document on pages two onward would sit above it and say the same thing
 * twice. Word's answer is a title page: `w:titlePg` plus a `first` header or
 * footer part, and it is one switch for both strips — once it is on, page one
 * takes both from the `first` parts and shows nothing where a part is missing.
 *
 * That last clause is the trap, and the reason this case exists. A document
 * that names only a first *header* must not silently lose its footer off page
 * one: absent means "page one is like every other page", which is the default
 * part repeated, not an empty one.
 */

const lines = Array.from({ length: 60 }, (_, index) => (
  <Paragraph id={`p${index}`}>
    {`Body line ${index + 1}, here to carry the document onto a second page so the ` +
      "first page and the rest have something to differ about."}
  </Paragraph>
));

export default defineCase({
  id: "furniture/first-page",
  feature: "furniture.firstPage",
  title: "A first page with furniture of its own",
  word: "Header & Footer → Different First Page (w:titlePg)",

  /**
   * Right in the file, right in Word, and only half-drawn in the preview.
   *
   * docx-preview renders one set of running furniture per section, and for a
   * title-page document it picks the first page's. The default header part is
   * in the package, Word draws it from page two onward, and it never reaches
   * the DOM at all — so there is nothing for the paginator to carry onto the
   * later sheets, and they show page one's letterhead instead of the running
   * strip.
   *
   * The paginator did not introduce this and cannot fix it: it can only
   * propagate the furniture docx-preview drew. Fixing it means rendering the
   * default header part, which docx-preview does not expose — and building the
   * markup from the model instead would be a second renderer for furniture,
   * which is the one thing this framework refuses to grow.
   *
   * Until then: a title-page document previews its first page faithfully and
   * shows the wrong strip above the rest. Word is right either way.
   */
  claim: "partial",
  knownRed: ["parity"],

  style: withBlocks({ strip: { spacingAfterPt: 0 } }),

  document: template(
    <Document
      id="first-page"
      title="First page furniture"
      firstHeader={
        <Paragraph id="fh" variant="strip">FERNHILL SYSTEMS — letterhead</Paragraph>
      }
      header={<Paragraph id="h" variant="strip">Continued · INV-2026-0142</Paragraph>}
      footer={<Paragraph id="f" variant="strip">Registered in England</Paragraph>}
    >
      {lines}
    </Document>
  ),

  regions: [
    { id: "first", anchor: "letterhead" },
    { id: "running", anchor: "Continued" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.section.titlePg, true, "the section is marked as having a title page");

      // Two header parts: the running one and page one's own.
      const headers = a.parts.filter((part) => /^word\/header\d+\.xml$/.test(part));
      is.equal(headers.length, 2, "the package holds two header parts");

      is.includes(a.documentXml, 'w:type="first"', "and the section points at the first-page one");
      is.includes(a.documentXml, 'w:type="default"', "as well as the running one");
    },

    word: (c, is) => {
      is.equal(c.differentFirstPage(), true, "Word reads it as a title page");

      is.equal(c.furniture("firstPage", "header").exists, true, "page one has a header of its own");
      is.includes(c.furniture("firstPage", "header").text, "letterhead", "carrying the letterhead");

      is.equal(c.furniture("primary", "header").exists, true, "and the other pages have the running one");
      is.includes(c.furniture("primary", "header").text, "Continued", "carrying the running text");

      // The trap: a document that named only a first *header* must not lose its
      // footer off page one. Absent means "like every other page".
      is.equal(c.furniture("firstPage", "footer").exists, true, "page one keeps a footer");
      is.includes(
        c.furniture("firstPage", "footer").text,
        "Registered in England",
        "and it is the running footer, repeated rather than emptied",
      );
    },

    preview: (b, is) => {
      is.includes(b.headerText(1), "letterhead", "the preview draws the letterhead on page one");
      is.equal(b.headerText(1).includes("Continued"), false, "and not the running strip as well");
    },

    parity: (p, is) => {
      is.equal(p.previewPages(), p.wordPages(), "both engines break the document into the same pages");
      is.includes(
        p.preview.headerText(2),
        "Continued",
        "and page two carries the running header, not the letterhead",
      );
    },
  },
});
