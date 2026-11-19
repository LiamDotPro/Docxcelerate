import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";

/**
 * A document that opens with Track Changes already on.
 *
 * It is a setting rather than content: `w:trackRevisions` tells Word to record
 * what is edited from here on. A freshly generated file has no revisions in it,
 * so the preview has nothing to draw and shows the text unchanged.
 */
export default defineCase({
  id: "document/track-changes",
  feature: "document.trackChanges",
  title: "Opens with Track Changes on",
  word: "Review → Track Changes (w:trackRevisions)",
  claim: "supported",

  document: template(
    <Document id="track-changes" title="Track changes" trackChanges>
      <Paragraph id="p">Either party may propose amendments to this agreement.</Paragraph>
    </Document>
  ),

  expect: {
    ooxml: (a, is) => {
      is.includes(a.settingsXml, "<w:trackRevisions", "the document turns Track Changes on");
      is.excludes(a.documentXml, "<w:ins", "and records no revisions of its own");
    },
    preview: (b, is) => {
      is.includes(b.para("Either party").text, "amendments", "the preview draws the text unchanged, with no revision marks");
    },
    word: (c, is) => {
      is.equal(c.trackRevisions(), true, "Word opens it with Track Changes on");
    },
  },
});
