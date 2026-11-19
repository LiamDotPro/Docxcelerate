import { Document, Graph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { withBlocks } from "../_support/style.ts";

/**
 * A theme sizes its charts' text.
 *
 * The `chart` block's `fontSizePt` letters the axes, key and axis titles, and
 * the title follows it three points larger. The Word probe does not read chart
 * font sizes, so its tier asserts only that Word accepted the part and laid a
 * plot out — a size outside what the schema allows is refused outright.
 */
const HEIGHT_PT = 200;
const TITLE = "Revenue by quarter";

export default defineCase({
  id: "charts/text-size",
  feature: "chart.textSize",
  title: "A theme's chart block sets the size of a chart's text",
  word: "Format Chart Area → Font size (c:txPr / a:defRPr@sz)",
  claim: "supported",

  style: withBlocks({ chart: { fontSizePt: 14 } }),

  document: template(
    <Document id="text-size" title="Charts">
      <Graph
        id="revenue"
        title={TITLE}
        graphType="bar"
        height={HEIGHT_PT}
        categoryAxisTitle="Quarter"
        valueAxisTitle="Revenue"
        data={{
          categories: ["Q1", "Q2", "Q3", "Q4"],
          series: [
            { label: "2024", values: [12, 18, 9, 22] },
            { label: "2025", values: [16, 14, 21, 27] },
          ],
        }}
      />
    </Document>
  ),

  expect: {
    ooxml: (a, is) => {
      const chart = a.chart(0);
      const xml = chart.xml ?? "";

      is.includes(xml, 'sz="1400" b="0" i="0"', "the chart's text is set at the theme's size");
      is.includes(xml, '<a:defRPr sz="1700" b="0">', "and its title three points larger");
      is.excludes(xml, 'sz="900"', "with nothing left at the default");
      is.excludes(xml, 'sz="1200"', "and no title at the default");
      is.equal(chart.title, TITLE, "the title is the document's");
    },

    preview: (b, is) => {
      is.equal(b.charts.length, 1, "the preview lays out one chart frame");
      is.within(b.chart(0).h, b.pt(HEIGHT_PT), "1mm", "at the depth the file gives it");
      is.equal(b.chart(0).plotted, true, "and draws the chart inside it");
    },

    word: (c, is) => {
      const chart = c.chart(0);

      is.equal(c.charts.length, 1, "Word finds one chart");
      is.equal(chart.title, TITLE, "titled as the document said");
      is.greater(chart.plotWidth, 0, "and lays out a plot inside it");
      is.greater(chart.plotHeight, 0, "with a height");
    },

    parity: (p, is) => {
      is.within(p.previewChartWidth(0), p.wordChartWidth(0), "1mm", "the frame is as wide on screen as in Word");
      is.within(p.previewChartHeight(0), p.wordChartHeight(0), "1mm", "and as deep");
    },
  },
});
