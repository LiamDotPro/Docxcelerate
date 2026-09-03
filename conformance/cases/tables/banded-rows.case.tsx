import { Cell, Document, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { withBlocks } from "../_support/style.ts";

/**
 * The zebra: every other body row tinted, counted by the renderer.
 *
 * Word has a mechanism for this — a table style with banding — and the packer
 * deliberately does not use it. The tint is written onto each cell as ordinary
 * shading instead, and the reason is that the alternative does not survive
 * publishing: a variant chosen by a `.map` index works in the preview and dies
 * the moment the map becomes a loop the engine walks, because a variant is a
 * static string and every row would then get whatever the build happened to
 * decide. Counting the rows as they are *drawn* is the only place the count is
 * still true. Which also means the banding is visible to anything that reads
 * the file, rather than living in a style the reader has to resolve.
 *
 * The other half of the rule is what a band is counted *from*. A zebra is a
 * reading aid for a column of like rows, and what says a table is one of those
 * is its header — a table without one is a layout, a letterhead or a totals
 * block, where striping tints whichever cells happen to be on an odd row. So
 * the count starts at the first body row and a table with no header is never
 * striped at all, which the second table here is for.
 */
const BAND = "F3F4F6";

export default defineCase({
  id: "tables/banded-rows",
  feature: "table.bandedRows",
  title: "Every other body row tinted, counted from the header",
  word: "Table Styles → Banded rows — written as per-cell w:shd, not as a w:tblStyle",
  claim: "supported",

  style: withBlocks({
    /** The band. Named `rowAlt`, which is where the renderer looks for it. */
    rowAlt: { fill: BAND },
  }),

  document: template(
    <Document id="banded-rows" title="Banded rows">
      <Table id="charges" columns={[{ width: "auto" }]}>
        <Row id="head" header>
          <Cell id="h1">Charges</Cell>
        </Row>
        <Row id="r1"><Cell id="c1">Body row one</Cell></Row>
        <Row id="r2"><Cell id="c2">Body row two</Cell></Row>
        <Row id="r3"><Cell id="c3">Body row three</Cell></Row>
        <Row id="r4"><Cell id="c4">Body row four</Cell></Row>
      </Table>

      <Table id="layout" columns={[{ width: "auto" }]}>
        <Row id="l1"><Cell id="d1">Headerless row one</Cell></Row>
        <Row id="l2"><Cell id="d2">Headerless row two</Cell></Row>
      </Table>
    </Document>
  ),

  regions: [
    { id: "one", anchor: "Body row one" },
    { id: "two", anchor: "Body row two" },
    { id: "four", anchor: "Body row four" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.cell("Body row one").shd, null, "the first body row is left plain");
      is.equal(a.cell("Body row two").shd, BAND, "the second is tinted");
      is.equal(a.cell("Body row three").shd, null, "the third plain again");
      is.equal(a.cell("Body row four").shd, BAND, "and the fourth tinted");

      // The claim about *how*: the banding is in the cells, not in a style the
      // reader has to resolve.
      is.equal(a.table(0).style, null, "leaning on no Word table style to do it");

      is.equal(a.cell("Headerless row one").shd, null, "a table with no header is not a column of like rows");
      is.equal(a.cell("Headerless row two").shd, null, "so nothing in it is striped");
    },

    preview: (b, is) => {
      is.equal(b.cell("Body row one").background, "rgba(0, 0, 0, 0)", "the preview leaves the first body row untinted");
      is.equal(b.cell("Body row two").background, b.hex(BAND), "and tints the second");
      is.equal(b.cell("Body row four").background, b.hex(BAND), "and the fourth");
      is.equal(b.cell("Headerless row two").background, "rgba(0, 0, 0, 0)", "and stripes nothing in the headerless table");
    },

    word: (c, is) => {
      is.equal(c.cell("Body row one").shading, null, "Word reads the first body row as unfilled");
      is.equal(c.cell("Body row two").shading, BAND, "the second as tinted");
      is.equal(c.cell("Body row four").shading, BAND, "and the fourth");
      is.equal(c.table(0).style, "Table Normal", "under no table style of its own");
    },
  },
});
