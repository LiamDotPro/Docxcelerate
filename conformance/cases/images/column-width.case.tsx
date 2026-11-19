import { Cell, Document, Image, Paragraph, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A picture told to be full is as wide as where it stands.
 *
 * On the page that is the text column; in a cell it is the cell's inside, the
 * column less its padding. What is measured is the extent the file states,
 * then that each engine drew the frame at it.
 */
// 4x2 PNG.
const pngWide = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAADwyuo0AAAAC0lEQVR42mNgwAUAABoAAS+Yl6YAAAAASUVORK5CYII=";

export default defineCase({
  id: "images/column-width",
  feature: "image.columnWidth",
  title: "A picture can fill the text column or its cell",
  word: "Insert → Pictures, resized to the text column (wp:extent cx/cy)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="column-width" title="Images">
      <Paragraph id="before">A paragraph above the pictures.</Paragraph>
      <Image id="hero" src={pngWide} alt="Hero" width="full" />
      <Table id="t" columns={[{ width: "auto" }, { width: "auto" }]}>
        <Row>
          <Cell><Image id="in-cell" src={pngWide} alt="In a cell" width="full" /></Cell>
          <Cell><Paragraph id="beside">Beside it.</Paragraph></Cell>
        </Row>
      </Table>
    </Document>
  ),

  regions: [],

  expect: {
    ooxml: (a, is) => {
      const xml = a.documentXml ?? "";

      is.includes(xml, '<wp:extent cx="6124575" cy="3057525"/>', "the picture is the text column wide");
      is.includes(xml, '<wp:extent cx="2857500" cy="1428750"/>', "and in a cell, the cell's inside");
    },

    preview: (b, is) => {
      const pictures = b.charts.filter((f) => f.picture);

      is.within(pictures[0].w, b.pt(481.89), "1mm", "the first fills the column");
      is.within(pictures[0].h, b.pt(240.94), "1mm", "and keeps its proportions");
      is.within(pictures[1].w, b.pt(224.95), "1mm", "the second fills its cell less the padding");
    },

    word: (c, is) => {
      is.within(c.shape(0).width, 481.89, 1, "Word reads the first at the column's width");
    },
  },
});
