import { Document, Graph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A value axis that does not start at zero.
 *
 * Left to itself Word scales the value axis from zero, so a series sitting
 * between 99.4 and 99.99 is drawn as one flat line near the top. The bounds
 * and the step are the fix, and they have to live in the chart part itself
 * (`c:scaling`, `c:majorUnit`) rather than in anything drawn over it, so that
 * the chart stays a real Word chart and Format Axis shows the author's numbers.
 *
 * Two charts hold the same figures: one bounded, one not. The unbounded one is
 * the control — it proves the bounds are written only where asked, and that a
 * chart with none is the file it always was.
 *
 * Only Word can say it honoured them, so the word tier reads the scale back
 * from the axis object rather than trusting the XML.
 */
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
const UPTIME = [{ label: "API", values: [99.62, 99.91, 99.48, 99.97, 99.85, 99.99] }];

const HEIGHT_PT = 140;

export default defineCase({
  id: "charts/value-axis-bounds",
  feature: "chart.valueAxisBounds",
  title: "A value axis fixed to 99 to 100, beside one left to Word",
  word: "Format Axis → Bounds / Units (c:scaling c:min c:max, c:majorUnit)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="value-axis-bounds" title="Value axis bounds">
      <Graph
        id="bounded"
        graphType="line"
        numberFormat="0.0"
        height={HEIGHT_PT}
        legend="none"
        valueAxisMin={99}
        valueAxisMax={100}
        valueAxisStep={0.2}
        data={{ categories: MONTHS, series: UPTIME }}
      />
      <Graph
        id="auto"
        graphType="line"
        numberFormat="0.0"
        height={HEIGHT_PT}
        legend="none"
        data={{ categories: MONTHS, series: UPTIME }}
      />
    </Document>,
  ),

  expect: {
    ooxml: (a, is) => {
      is.equal(a.chartCount, 2, "every chart is a part of its own");

      is.includes(
        a.chart(0).xml ?? "",
        '<c:max val="100"/><c:min val="99"/>',
        "the bounds are written, max before min, as the schema orders them",
      );
      is.includes(a.chart(0).xml ?? "", '<c:majorUnit val="0.2"/>', "and so is the step");
      is.excludes(a.chart(1).xml ?? "", "<c:min ", "a chart given no bounds writes none");
      is.excludes(a.chart(1).xml ?? "", "<c:majorUnit", "nor a step");
    },

    preview: (b, is) => {
      is.equal(b.charts.length, 2, "the preview lays out every chart's frame");
      is.equal(
        b.charts.every((chart) => chart.plotted),
        true,
        "and draws a plot in every one",
      );
    },

    word: (c, is) => {
      is.equal(c.charts.length, 2, "Word opens the file and builds both charts");
      is.equal(c.chart(0).valueMin, 99, "Word reads the lower bound");
      is.equal(c.chart(0).valueMax, 100, "the upper bound");
      is.equal(c.chart(0).valueStep, 0.2, "and the step");
    },

    parity: (p, is) => {
      is.within(p.previewChartWidth(0), p.wordChartWidth(0), "1mm", "the frame agrees");
      is.within(p.previewChartHeight(0), p.wordChartHeight(0), "1mm", "in both directions");
    },
  },
});
