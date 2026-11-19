import { Document, Graph, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle, COLUMN_MM } from "../_support/style.ts";

/**
 * A chart with its figures printed under the plot.
 *
 * The table is part of the chart part, not a second object beside it, so the
 * numbers move with the chart and reflow when its data is edited. Its rows
 * carry the series' keys, which is why the chart has no legend of its own.
 * `c:dTable` has a fixed place in the plot area — after the axes, before the
 * plot's fill — and a table anywhere else is a file Word offers to repair, so
 * the ooxml tier pins the neighbours.
 */
const QUARTERS = ["Q1", "Q2", "Q3", "Q4"];
const LAST_YEAR = [12400, 18100, 9300, 22800];
const THIS_YEAR = [16200, 14500, 21700, 27300];

const HEIGHT_PT = 240;

export default defineCase({
  id: "charts/data-table",
  feature: "chart.dataTable",
  title: "A column chart with a data table and legend keys under the plot",
  word: "Chart Design → Add Chart Element → Data Table → With Legend Keys (c:dTable)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="data-table" title="Charts">
      <Graph
        id="revenue"
        graphType="bar"
        dataTable
        numberFormat="#,##0"
        height={HEIGHT_PT}
        data={{
          categories: QUARTERS,
          series: [
            { label: "2024", values: LAST_YEAR },
            { label: "2025", values: THIS_YEAR },
          ],
        }}
      />
      <Paragraph id="caption">Figure 1: revenue by quarter.</Paragraph>
    </Document>
  ),

  regions: [{ id: "revenue", anchor: "Figure 1" }],

  expect: {
    ooxml: (a, is) => {
      const chart = a.chart(0);

      is.equal(chart.plot, "bar", "the plot is a bar chart");
      is.includes(chart.xml ?? "", "<c:dTable>", "the plot area holds a data table");
      is.includes(chart.xml ?? "", '<c:showKeys val="1"/>', "with the legend keys");
      is.includes(chart.xml ?? "", "</c:dTable><c:spPr>", "placed before the plot's fill");
      is.equal(chart.legend, null, "and no legend, the table carries the keys");
      is.equal(a.chartSeries(0, 0).values, LAST_YEAR, "the first series' figures are cached");
      is.equal(a.chartSeries(0, 1).values, THIS_YEAR, "and the second's");
    },

    preview: (b, is) => {
      const chart = b.chart(0);

      is.equal(b.charts.length, 1, "the preview lays out one chart frame");
      is.within(chart.w, b.mm(COLUMN_MM), "1mm", "as wide as the text column");
      is.within(chart.h, b.pt(HEIGHT_PT), "1mm", "and as deep as the file says, table included");
      is.equal(chart.plotted, true, "and draws the chart inside it");
    },

    word: (c, is) => {
      const chart = c.chart(0);

      is.equal(c.charts.length, 1, "Word finds one chart");
      is.equal(chart.typeName, "columnClustered", "and builds a clustered column from it");
      is.equal(chart.hasLegend, false, "with no legend of its own");
      is.equal(c.chartSeries(0, 0).values, LAST_YEAR, "Word reads the first series back");
      is.equal(c.chartSeries(0, 1).values, THIS_YEAR, "and the second");
    },

    parity: (p, is) => {
      is.within(p.previewChartWidth(0), p.wordChartWidth(0), "1mm", "the frame is as wide on screen as in Word");
      is.within(p.previewChartHeight(0), p.wordChartHeight(0), "1mm", "and as deep");
    },
  },
});
