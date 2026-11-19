import { Document, Graph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A trendline, packed as one rather than as figures.
 *
 * The line is a `c:trendline` on the series, so Word works it out from the
 * cached values and works it out again when a reader uses Edit Data. The
 * alternative, a second series of hand-computed numbers, is right on the day it
 * is written and stale the first time anyone edits the data.
 *
 * That is also what the assertions hold to: the trendline adds no cached
 * figures, it is not a series, and the series it sits on keeps its own colour.
 */
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
const SIGNUPS = [120, 135, 128, 160, 171, 190];
const CHURN = [30, 42, 25, 38, 29, 41];

/** The first colour of the shipped chart palette. */
const SERIES_ONE = "2A78D6";

const HEIGHT_PT = 90;

export default defineCase({
  id: "charts/trendline",
  feature: "chart.trendline",
  title: "A linear trendline and a moving average, drawn by Word from the data",
  word: "Chart Elements → Trendline (c:trendline)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="trendline" title="Trendline">
      <Graph
        id="linear"
        graphType="line"
        height={HEIGHT_PT}
        legend="none"
        data={{
          categories: MONTHS,
          series: [
            { label: "Sign-ups", values: SIGNUPS, trendline: "linear" },
            { label: "Churn", values: CHURN },
          ],
        }}
      />
      <Graph
        id="average"
        graphType="bar"
        height={HEIGHT_PT}
        legend="none"
        data={{ categories: MONTHS, series: [{ label: "Churn", values: CHURN, trendline: { movingAverage: 3 } }] }}
      />
    </Document>
  ),

  expect: {
    ooxml: (a, is) => {
      is.equal(a.chartCount, 2, "every chart is a part of its own");
      is.includes(a.chart(0).xml ?? "", `<c:trendlineType val="linear"/>`, "the line is a native trendline");
      is.equal((a.chart(0).xml ?? "").split("<c:trendline>").length - 1, 1, "on the series that asked and no other");
      is.includes(
        a.chart(1).xml ?? "",
        `<c:trendlineType val="movingAvg"/><c:period val="3"/>`,
        "a moving average carries its window",
      );
      is.equal(a.chartSeries(0, 0).values, SIGNUPS, "the trendline adds no cached figures");
      is.equal(a.chartSeries(1, 0).values, CHURN, "to either chart");
      is.equal(a.chartSeries(0, 0).color, SERIES_ONE, "and the series keeps the palette's colour, not the line's");
    },

    preview: (b, is) => {
      is.equal(b.charts.length, 2, "the preview lays out every chart's frame");
      is.within(b.chart(0).h, b.pt(HEIGHT_PT), "1mm", "each at the depth the file gives it");
      is.equal(b.charts.every((chart) => chart.plotted), true, "and draws a plot in every one");
    },

    word: (c, is) => {
      is.equal(c.charts.length, 2, "Word opens the file and builds every chart");
      is.equal(c.chart(0).seriesCount, 2, "a trendline is not a series");
    },

    parity: (p, is) => {
      is.within(p.previewChartWidth(0), p.wordChartWidth(0), "1mm", "the first frame agrees");
      is.within(p.previewChartHeight(0), p.wordChartHeight(0), "1mm", "in both directions");
    },
  },
});
