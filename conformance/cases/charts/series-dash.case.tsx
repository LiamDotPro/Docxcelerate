import { Document, Graph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A forecast and a target, told apart from the actuals by their line.
 *
 * Colour alone is a poor way to say "this one is not like the others": a
 * greyscale print flattens it and a colour-blind reader may never see it. A
 * dash is the stroke's own style, so it survives both. The packer writes it as
 * `a:prstDash` on the series line, which is what Word's Format Data Series →
 * Line → Dash type reads and what the key beside the plot repeats.
 *
 * What is asserted is that the preset is in the file exactly twice — once per
 * dashed series — so the solid actual and every marker outline carry none.
 * The preview tier checks only that the chart still draws; the word tier that
 * Word opens a file with the preset in place.
 */
const QUARTERS = ["Q1", "Q2", "Q3", "Q4"];
const FORECAST = [12, 17, 22, 27];

const HEIGHT_PT = 200;

export default defineCase({
  id: "charts/series-dash",
  feature: "chart.seriesDash",
  title: "A forecast dashed and a target dotted against a solid actual",
  word: "Format Data Series → Line → Dash type (a:prstDash)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="series-dash" title="Charts">
      <Graph
        id="revenue"
        graphType="line"
        height={HEIGHT_PT}
        data={{
          categories: QUARTERS,
          series: [
            { label: "Actual", values: [12, 18, 21, null] },
            { label: "Forecast", values: FORECAST, dash: "dashed" },
            { label: "Target", values: [20, 20, 20, 20], dash: "dotted" },
          ],
        }}
      />
    </Document>
  ),

  expect: {
    ooxml: (a, is) => {
      const xml = a.chart(0).xml ?? "";

      is.equal(a.chart(0).plot, "line", "the plot is a line chart");
      is.includes(xml, '<a:prstDash val="dash"/>', "the forecast is dashed");
      is.includes(xml, '<a:prstDash val="sysDot"/>', "the target is dotted");
      is.equal(xml.match(/<a:prstDash /g)?.length, 2, "the actual stays solid");
      is.equal(a.chartSeries(0, 1).values, FORECAST, "and the dashed series keeps its numbers");
    },

    preview: (b, is) => {
      is.equal(b.charts.length, 1, "the preview lays out one chart frame");
      is.within(b.chart(0).h, b.pt(HEIGHT_PT), "1mm", "as deep as the file says");
      is.equal(b.chart(0).plotted, true, "and draws the chart inside it");
    },

    word: (c, is) => {
      is.equal(c.charts.length, 1, "Word finds one chart");
      is.equal(c.chart(0).typeName, "lineMarkers", "and builds a line chart with markers");
      is.equal(c.chart(0).seriesCount, 3, "with all three series");
    },
  },
});
