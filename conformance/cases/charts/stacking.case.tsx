import { Document, Graph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * Stacking, and the two different things it can mean.
 *
 * `stacked` is a property rather than a chart type, so one flag has to reach
 * Word as four different charts — a stacked column, a hundred per cent stacked
 * column, and the same again lying down, on a line and on an area. Each is its
 * own constant in Word, and each is a plausible way for the grouping to be
 * written and quietly ignored.
 *
 * **A percent stack is not a stack with a different label.** Word works each
 * column's shares out from the raw figures itself: the file carries the same
 * numbers either way, and it is `c:grouping` alone that decides whether the
 * reader sees totals or shares. So the assertion that matters is that the
 * grouping arrived, not that the values did — a chart whose grouping was
 * dropped still holds every figure and still draws, as the wrong chart.
 *
 * The axis is the other half. A percent stack's scale runs from zero to one,
 * and printed unformatted Word labels it 0, 0.2, 0.4 — which is not what
 * anybody means by a 100% stacked chart. The packer writes `0%` where the
 * document named no format of its own, and `sourceLinked="0"` is what stops
 * the numbers' own format taking the axis back.
 *
 * The types are Word's own constants: 52 column stacked, 53 column stacked to
 * 100%, 59 bar stacked to 100%, 67 line with markers stacked to 100%, 77 area
 * stacked to 100%.
 */
const MONTHS = ["Jan", "Feb", "Mar"];
const SERIES = [
  { label: "North", values: [12, 18, 9] },
  { label: "South", values: [16, 14, 21] },
];

const HEIGHT_PT = 90;

/** Each way of stacking, and the chart Word should build from it. */
const STACKS = [
  { id: "column-total", graphType: "bar", stacked: true, grouping: "stacked", word: "columnStacked" },
  {
    id: "column-share",
    graphType: "bar",
    stacked: "percent",
    grouping: "percentStacked",
    word: "columnStacked100",
  },
  {
    id: "bar-share",
    graphType: "barHorizontal",
    stacked: "percent",
    grouping: "percentStacked",
    word: "barStacked100",
  },
  {
    id: "line-share",
    graphType: "line",
    stacked: "percent",
    grouping: "percentStacked",
    word: "lineMarkersStacked100",
  },
  {
    id: "area-share",
    graphType: "area",
    stacked: "percent",
    grouping: "percentStacked",
    word: "areaStacked100",
  },
] as const;

export default defineCase({
  id: "charts/stacking",
  feature: "chart.stacking",
  title: "Stacked to a total and stacked to a share, on four kinds of plot",
  word: "Insert → Chart → Stacked / 100% Stacked (c:grouping)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="stacking" title="Stacking">
      {STACKS.map((stack) => (
        <Graph
          id={stack.id}
          graphType={stack.graphType}
          stacked={stack.stacked}
          height={HEIGHT_PT}
          legend="none"
          data={{ categories: MONTHS, series: SERIES }}
        />
      ))}
    </Document>
  ),

  expect: {
    ooxml: (a, is) => {
      is.equal(a.chartCount, STACKS.length, "every chart is a part of its own");

      STACKS.forEach((stack, index) => {
        is.equal(
          a.chart(index).grouping,
          stack.grouping,
          `${stack.id} is grouped as ${stack.grouping}`,
        );
      });

      // A stack whose segments do not close up draws as a staircase, which is
      // the one failure here that looks like a chart rather than like nothing.
      is.includes(a.chart(0).xml ?? "", `<c:overlap val="100"/>`, "a stack closes its segments up");
      is.includes(a.chart(1).xml ?? "", `<c:overlap val="100"/>`, "and so does a percent stack");

      // The axis reads as a share. Left to the numbers' own format it would
      // run 0, 0.2, 0.4 — a scale nobody means by "100% stacked".
      is.includes(
        a.chart(1).xml ?? "",
        `<c:numFmt formatCode="0%" sourceLinked="0"/>`,
        "a percent stack labels its value axis as a share",
      );
      is.excludes(
        a.chart(0).xml ?? "",
        `formatCode="0%"`,
        "while a stack to a total keeps the numbers' own format",
      );

      // The figures are the raw ones on both. Word divides; the file does not,
      // and a packer that pre-divided would hand Word shares of shares.
      is.equal(a.chartSeries(1, 0).values, SERIES[0].values, "the file carries the raw figures");
      is.equal(a.chartSeries(1, 1).values, SERIES[1].values, "for every series");
    },

    preview: (b, is) => {
      is.equal(b.charts.length, STACKS.length, "the preview lays out every chart's frame");
      is.within(b.chart(0).h, b.pt(HEIGHT_PT), "1mm", "each at the depth the file gives it");
      is.equal(
        b.charts.every((chart) => chart.plotted),
        true,
        "and draws a plot in every one",
      );
    },

    word: (c, is) => {
      is.equal(c.charts.length, STACKS.length, "Word opens the file and builds every chart");

      STACKS.forEach((stack, index) => {
        is.equal(
          c.chart(index).typeName,
          stack.word,
          `Word makes ${stack.id} a ${stack.word}`,
        );
      });

      // The distinction the whole case is about, stated once as Word sees it:
      // the same numbers, the same plot, two different charts.
      is.equal(c.chart(0).typeName, "columnStacked", "a total and a share are not the same chart");
      is.equal(c.chart(1).typeName, "columnStacked100", "even drawn from the same figures");
    },

    parity: (p, is) => {
      is.within(p.previewChartWidth(0), p.wordChartWidth(0), "1mm", "the first frame agrees");
      is.within(p.previewChartHeight(0), p.wordChartHeight(0), "1mm", "in both directions");
    },
  },
});
