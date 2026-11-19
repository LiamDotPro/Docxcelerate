import { Document, Paragraph, Section, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { withBlocks } from "../_support/style.ts";

/**
 * A section heading set apart by its variant.
 *
 * `variant` reached a section's node and was dropped by the renderer, so every
 * heading in a document looked alike and a theme could not set a schedule or an
 * annex apart from a clause. The variant now restyles the heading as direct
 * formatting over Heading 1, so the heading is still what the contents and
 * Word's outline find, and a clause beside it is left exactly as it was.
 */
export default defineCase({
  id: "text/section-variant",
  feature: "section.variant",
  title: "A section heading set by its variant",
  word: "Heading 1 + direct formatting (w:jc, w:caps, w:pBdr, w:spacing, w:keepNext)",
  claim: "supported",

  style: withBlocks({
    schedule: {
      align: "center",
      transform: "uppercase",
      fontSizePt: 14,
      border: "111827",
      borderSides: ["bottom"],
      spacingBeforePt: 24,
      keepWithNext: true,
    },
  }),

  document: template(
    <Document id="section-variant" title="Section variant">
      <Section id="clause" title="1. Definitions">
        <Paragraph id="a">A clause keeps the theme's heading.</Paragraph>
      </Section>
      <Section id="schedule" title="Schedule 1: Services" variant="schedule">
        <Paragraph id="b">A schedule sets its own.</Paragraph>
      </Section>
    </Document>
  ),

  regions: [
    { id: "clause", anchor: "1. Definitions" },
    { id: "schedule", anchor: "Schedule 1" },
  ],

  expect: {
    ooxml: (a, is) => {
      const heading = a.para("Schedule 1");
      is.equal(heading.style, "Heading1", "it is still a Heading 1");
      is.equal(heading.jc, "center", "centred");
      is.equal(heading.run.caps, true, "in capitals");
      is.equal(heading.run.szHalfPt, 28, "at 14pt");
      is.equal(heading.keepNext, true, "kept with what follows");
      is.equal(heading.spacing.before, 480, "with 24pt above");
      is.equal(heading.pBdr.bottom.color, "111827", "ruled underneath");
      is.equal(heading.pBdr.top, null, "and only underneath");
      is.equal(a.para("1. Definitions").jc, null, "the clause heading is not centred");
      is.equal(a.para("1. Definitions").run.caps, null, "nor in capitals");
    },

    preview: (b, is) => {
      const heading = b.para("Schedule 1");
      is.equal(heading.textAlign, "center", "the preview centres it");
      is.equal(heading.textTransform, "uppercase", "in capitals");
      is.within(heading.fontSize, b.pt(14), 0.5, "at 14pt");
      is.equal(heading.borderBottomColor, b.hex("111827"), "with the rule");
      is.notEqual(b.para("1. Definitions").textAlign, "center", "the clause heading is not centred");
    },

    word: (c, is) => {
      is.equal(c.para("Schedule 1").alignment, "center", "Word reads the schedule heading as centred");
    },

    parity: (p, is) => {
      is.within(p.previewX("Schedule 1"), p.wordX("Schedule 1"), "1mm", "the heading starts where Word starts it");
    },
  },
});
