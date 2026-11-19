import { Cell, Document, Graph, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * Charts without axes or gridlines: the sparkline.
 *
 * `axes={false}` writes both axes with `c:delete val="1"` rather than leaving
 * them out. Every plot group names two `c:axId`s, so a part whose axes are
 * missing is one Word will not open; a deleted axis still scales the plot but
 * draws no line and no labels, and the plot takes the frame.
 *
 * The grid is independent: `gridlines={false}` alone keeps the axes.
 */
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May"];
const TREND = [12, 18, 9, 22, 27];
const MIX = [16, 14, 21, 11, 19];

export default defineCase({
  id: "charts/hidden-axes",
  feature: "chart.hiddenAxes",
  title: "A sparkline in a table cell, and a chart with its grid dropped",
  word: "Chart Elements → Axes / Gridlines unchecked (c:delete, c:majorGridlines)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="hidden-axes" title="Hidden axes">
      <Table id="kpis" columns={[{ width: "auto" }, { width: 34 }]}>
        <Row id="revenue-row">
          <Cell id="revenue-figure">Revenue 1.2m</Cell>
          <Cell id="revenue-trend">
            <Graph
              id="spark"
              graphType="line"
              axes={false}
              gridlines={false}
              width={72}
              height={24}
              legend="none"
              data={{ categories: MONTHS, series: [{ values: TREND }] }}
            />
          </Cell>
        </Row>
      </Table>
      <Graph
        id="ungridded"
        graphType="bar"
        gridlines={false}
        height={90}
        legend="none"
        data={{ categories: MONTHS, series: [{ values: MIX }] }}
      />
    </Document>
  ),

  expect: {
    ooxml: (a, is) => {
      is.equal(a.chartCount, 2, "both charts are parts of their own");
      is.includes(a.chart(0).xml ?? "", `<c:delete val="1"/>`, "the sparkline's axes are deleted");
      is.excludes(a.chart(0).xml ?? "", `<c:delete val="0"/>`, "none of them drawn");
      is.excludes(a.chart(0).xml ?? "", `<c:majorGridlines>`, "and no grid");
      is.equal(a.chart(0).axisCount, 2, "deleted, not dropped");
      is.includes(a.chart(1).xml ?? "", `<c:delete val="0"/>`, "the ungridded chart keeps its axes");
      is.excludes(a.chart(1).xml ?? "", `<c:majorGridlines>`, "without a grid");
      is.equal(a.chartSeries(0, 0).values, TREND, "the figures are still carried");
    },

    preview: (b, is) => {
      is.equal(b.charts.length, 2, "the preview lays out both frames");
      is.within(b.chart(0).w, b.pt(72), "1mm", "the sparkline at its width");
      is.within(b.chart(0).h, b.pt(24), "1mm", "and its height");
      is.equal(b.charts.every((chart) => chart.plotted), true, "and draws a plot in each");
    },

    word: (c, is) => {
      is.equal(c.charts.length, 2, "Word opens the file and builds both charts");
      is.equal(c.chart(0).typeName, "lineMarkers", "the sparkline is a line");
      is.equal(c.chart(1).typeName, "columnClustered", "the other a column chart");
    },

    parity: (p, is) => {
      is.within(p.previewChartWidth(0), p.wordChartWidth(0), "1mm", "the sparkline's frame agrees");
      is.within(p.previewChartHeight(0), p.wordChartHeight(0), "1mm", "in both directions");
    },
  },
});
