import { Document, Graph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * Every kind of chart the toolkit can ask Word for.
 *
 * `charts/column-series` proves the chain for one chart. This proves that the
 * choice of plot reaches Word intact for all of them, which is a different
 * failure: a chart group written with its children in the wrong order, or with
 * an axis a type does not take, gives a package Word refuses to open outright
 * — not a chart drawn wrongly, but a document that will not open at all. That
 * failure is invisible to probe A, because the XML it wrote is exactly the XML
 * it meant to write. Only Word can say whether the schema was obeyed.
 *
 * Measured while writing this: a pie declaring axes, or a stacked bar left at
 * the clustered overlap, or a scatter given a category axis instead of a
 * second value axis — each is a plausible-looking file, and Word's answer to
 * all three is the same "Word experienced an error trying to open the file".
 * So the assertion that matters most here is the dullest one: that Word opened
 * it and found seven charts.
 *
 * The types are Word's own constants, named rather than quoted at a reader:
 * 51 clustered column, 57 clustered bar, 65 line with markers, 1 area, 5 pie,
 * -4120 doughnut, 74 scatter with lines.
 */
const MONTHS = ["Jan", "Feb", "Mar", "Apr"];
const SERIES = [
  { label: "North", values: [12, 18, 9, 22] },
  { label: "South", values: [16, 14, 21, 27] },
];
const ONE_SERIES = [{ label: "Share", values: [42, 28, 18, 12] }];

const HEIGHT_PT = 90;

/** Each kind, and the chart Word should build from it. */
const KINDS = [
  { graphType: "bar", plot: "bar", word: "columnClustered" },
  { graphType: "barHorizontal", plot: "bar", word: "barClustered" },
  { graphType: "line", plot: "line", word: "lineMarkers" },
  { graphType: "area", plot: "area", word: "area" },
  { graphType: "pie", plot: "pie", word: "pie" },
  { graphType: "doughnut", plot: "doughnut", word: "doughnut" },
  { graphType: "scatter", plot: "scatter", word: "xyScatterLines" },
] as const;

export default defineCase({
  id: "charts/kinds",
  feature: "chart.kinds",
  title: "Bar, column, line, area, pie, doughnut and scatter, each as its own plot",
  word: "Insert → Chart → every type (c:barChart, c:lineChart, c:pieChart …)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="kinds" title="Chart kinds">
      {KINDS.map((kind) => (
        <Graph
          id={kind.graphType}
          graphType={kind.graphType}
          height={HEIGHT_PT}
          legend="none"
          data={{
            categories: MONTHS,
            series: kind.plot === "pie" || kind.plot === "doughnut" ? ONE_SERIES : SERIES,
          }}
        />
      ))}
    </Document>
  ),

  expect: {
    ooxml: (a, is) => {
      is.equal(a.chartCount, KINDS.length, "every chart is a part of its own");

      KINDS.forEach((kind, index) => {
        is.equal(a.chart(index).plot, kind.plot, `${kind.graphType} packs as a ${kind.plot} chart`);
        is.equal(a.chart(index).partPresent, true, `and ${kind.graphType}'s part is in the package`);
      });

      // The two that share a plot element and differ only in direction.
      is.equal(a.chart(0).barDir, "col", "a bar chart stands its bars up");
      is.equal(a.chart(1).barDir, "bar", "and a horizontal one lays them down");

      // A horizontal bar reads in the order its categories were written, which
      // OOXML's default does not: `minMax` puts the first category at the
      // bottom, the way Word's own bar charts come out. The scale is reversed
      // and the value axis told to cross at the far end, so the labels stay on
      // the left where they were.
      is.includes(
        a.chart(1).xml ?? "",
        `<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="maxMin"/>`,
        "a horizontal bar reverses its category scale",
      );
      is.includes(a.chart(1).xml ?? "", `<c:crosses val="max"/>`, "and the value axis crosses at the far end");
      is.includes(
        a.chart(0).xml ?? "",
        `<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/>`,
        "while a standing bar does not",
      );

      // A pie has no axes to have; every other kind has two. A pie that
      // declared them is a file Word will not open.
      is.equal(a.chart(4).axisCount, 0, "a pie declares no axes");
      is.equal(a.chart(5).axisCount, 0, "and neither does a doughnut");
      is.equal(a.chart(0).axisCount, 2, "a bar declares two");
      is.equal(a.chart(6).axisCount, 2, "and so does a scatter");
      is.includes(a.chart(6).xml ?? "", "<c:xVal>", "whose x values are numbers rather than names");

      // Every kind carries its data, whichever element the plot turned out to
      // be. A chart type that packed its series into the wrong child would
      // still be the right chart with nothing on it.
      KINDS.forEach((kind, index) => {
        is.greater(
          a.chartSeries(index, 0).values.length,
          0,
          `${kind.graphType} carries its values`,
        );
      });
    },

    preview: (b, is) => {
      is.equal(b.charts.length, KINDS.length, "the preview lays out every chart's frame");
      is.within(b.chart(0).h, b.pt(HEIGHT_PT), "1mm", "each at the depth the file gives it");
      is.equal(
        b.charts.every((chart) => chart.plotted),
        true,
        "and draws a plot in every one",
      );
    },

    word: (c, is) => {
      // The dull assertion that is the whole point: Word opened the file. A
      // chart group whose children are in the wrong order does not draw badly,
      // it makes the package unreadable, and this is where that shows up.
      is.equal(c.charts.length, KINDS.length, "Word opens the file and builds every chart");

      KINDS.forEach((kind, index) => {
        is.equal(
          c.chart(index).typeName,
          kind.word,
          `Word makes ${kind.graphType} a ${kind.word}`,
        );
      });

      is.equal(c.chartSeries(0, 0).values, SERIES[0].values, "and reads the figures back out of the first");
      is.equal(c.chartSeries(4, 0).values, ONE_SERIES[0].values, "and out of the pie");
    },

    parity: (p, is) => {
      is.within(p.previewChartWidth(0), p.wordChartWidth(0), "1mm", "the first frame agrees");
      is.within(p.previewChartHeight(0), p.wordChartHeight(0), "1mm", "in both directions");
    },
  },
});
