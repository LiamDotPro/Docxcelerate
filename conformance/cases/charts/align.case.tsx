import { Document, Graph, Image, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle, COLUMN_MM } from "../_support/style.ts";

/**
 * A chart or a picture, placed across the column.
 *
 * Both used to sit flush left with no way to say otherwise short of a
 * single-cell table. `align` is the paragraph's own `w:jc`, on the line the
 * drawing is in, so the assertion is on that line in the file and then on where
 * the drawing actually lands: the preview reads `w:jc` back as `text-align`,
 * and whether that moves a picture span and a chart frame is what tier B shows.
 */
const PIXEL = "data:image/png;base64," +
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

export default defineCase({
  id: "charts/align",
  feature: "picture.align",
  title: "A chart centred and a picture ranged right",
  word: "Paragraph → Alignment (w:jc) on a chart or picture",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="align" title="Align">
      <Paragraph id="before">A paragraph above.</Paragraph>

      <Graph
        id="centred"
        align="center"
        width={240}
        height={140}
        caption="Figure 1: centred."
        data={{ categories: ["Q1", "Q2"], series: [{ label: "2025", values: [4, 7] }] }}
      />

      <Image id="mark" align="right" src={PIXEL} alt="Mark" width={72} height={36} />

      <Paragraph id="after">A paragraph below.</Paragraph>
    </Document>
  ),

  expect: {
    ooxml: (a, is) => {
      is.equal(a.para("A paragraph above").jc, null, "an unaligned paragraph says nothing");
      is.equal(a.paraAt(1).jc, "center", "the chart's line is centred");
      is.equal(a.para("Figure 1").jc, "center", "and so is its caption");
      is.equal(a.paraAt(3).jc, "right", "the picture's line is ranged right");
    },

    preview: (b, is) => {
      const chart = b.chart(0);
      const picture = b.chart(1);

      is.within(chart.x + chart.w / 2, b.mm(COLUMN_MM) / 2, "2mm", "the chart sits in the middle of the column");
      is.equal(picture.picture, true, "the second frame is the picture");
      is.within(picture.x + picture.w, b.mm(COLUMN_MM), "2mm", "and its right edge is the column's");
    },

    word: (c, is) => {
      is.equal(c.paraAt(1).alignment, "center", "Word centres the chart's line");
      is.equal(c.paraAt(3).alignment, "right", "and ranges the picture's right");
    },
  },
});
