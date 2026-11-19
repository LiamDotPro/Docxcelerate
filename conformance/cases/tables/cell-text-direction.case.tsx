import { Cell, Document, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";

/**
 * Words turned on their side in a cell, for a heading over a column too narrow
 * to hold it across. `w:textDirection` is what says which way they read; the
 * row growing to fit them is Word's own doing, so the case checks the turned
 * heading ends up taller than it is wide rather than a particular depth.
 */
export default defineCase({
  id: "tables/cell-text-direction",
  feature: "table.cellTextDirection",
  title: "Headings turned up and down over narrow columns",
  word: "Table Properties → Cell → Text Direction (w:textDirection)",
  claim: "supported",

  document: template(
    <Document id="cell-text-direction" title="Cell text direction">
      <Table
        id="rota"
        columns={[{ width: 40 }, { width: 12 }, { width: 12 }, { width: "auto" }]}
      >
        <Row id="head" header>
          <Cell id="h1">Name</Cell>
          <Cell id="h2" textDirection="up">Monday early</Cell>
          <Cell id="h3" textDirection="down">Tuesday early</Cell>
          <Cell id="h4">Notes</Cell>
        </Row>
        <Row id="line">
          <Cell id="d1">Avery</Cell>
          <Cell id="d2">x</Cell>
          <Cell id="d3">x</Cell>
          <Cell id="d4">Covers the late</Cell>
        </Row>
      </Table>
    </Document>
  ),

  regions: [
    { id: "up", anchor: "Monday early" },
    { id: "down", anchor: "Tuesday early" },
    { id: "across", anchor: "Name" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.cell("Monday early").textDirection, "btLr", "up reads bottom to top");
      is.equal(a.cell("Tuesday early").textDirection, "tbRl", "down reads top to bottom");
      is.equal(a.cell("Name").textDirection, null, "a plain cell says nothing");
    },

    preview: (b, is) => {
      is.equal(b.cell("Monday early").writingMode, "vertical-rl", "the preview turns the first");
      is.notEqual(b.cell("Monday early").transform, "none", "and flips it to read upward");
      is.equal(b.cell("Tuesday early").writingMode, "vertical-rl", "turns the second");
      is.equal(b.cell("Tuesday early").transform, "none", "and leaves it reading downward");
      is.equal(b.cell("Name").writingMode, "horizontal-tb", "a plain cell stays across");
      is.greater(b.cell("Monday early").h, b.cell("Monday early").w, "the turned heading is taller than it is wide");
    },
  },
});
