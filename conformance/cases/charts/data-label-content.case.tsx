import { Document, Graph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * What a data label prints, beyond the figure.
 *
 * The category flag used to be hard-coded off, so a pie could never name its
 * slices and needed a key. The two-part labels also bring a `c:separator`,
 * which the schema wants after `c:showBubbleSize`: written anywhere else it is
 * a file Word offers to repair, so Word building every chart is the check.
 */
const CATEGORIES = ["Rent", "Food", "Travel"];
const SERIES = [{ label: "Spend", values: [1200, 900, 760] }];

const HEIGHT_PT = 90;

/** Each kind of label, and the flags the file should carry for it. */
const LABELS = [
  { id: "pie-name-share", graphType: "pie", dataLabels: "categoryAndPercent", flags: ["0", "1", "1"], separator: true },
  { id: "doughnut-name-value", graphType: "doughnut", dataLabels: "categoryAndValue", flags: ["1", "1", "0"], separator: true },
  { id: "bar-name", graphType: "bar", dataLabels: "category", flags: ["0", "1", "0"], separator: false },
  { id: "bar-value", graphType: "bar", dataLabels: true, flags: ["1", "0", "0"], separator: false },
] as const;

export default defineCase({
  id: "charts/data-label-content",
  feature: "chart.dataLabels",
  title: "Data labels that print a name, a value or a share",
  word: "Chart Elements → Data Labels → More Options (c:dLbls)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="data-label-content" title="Data label content">
      {LABELS.map((label) => (
        <Graph
          id={label.id}
          graphType={label.graphType}
          dataLabels={label.dataLabels}
          height={HEIGHT_PT}
          legend="none"
          data={{ categories: CATEGORIES, series: SERIES }}
        />
      ))}
    </Document>
  ),

  expect: {
    ooxml: (a, is) => {
      is.equal(a.chartCount, LABELS.length, "every chart is a part of its own");

      LABELS.forEach((label, index) => {
        const xml = a.chart(index).xml ?? "";
        const [val, cat, percent] = label.flags;

        is.includes(xml, `<c:showVal val="${val}"/>`, `${label.id} prints its value as asked`);
        is.includes(xml, `<c:showCatName val="${cat}"/>`, `${label.id} prints its category as asked`);
        is.includes(xml, `<c:showPercent val="${percent}"/>`, `${label.id} prints its share as asked`);

        if (label.separator) {
          is.includes(xml, `<c:separator>, </c:separator>`, `${label.id} joins its parts with a comma`);
        } else {
          is.excludes(xml, `<c:separator>`, `${label.id} has one part, so nothing to join`);
        }
      });
    },

    preview: (b, is) => {
      is.equal(b.charts.length, LABELS.length, "the preview lays out every chart's frame");
      is.equal(
        b.charts.every((chart) => chart.plotted),
        true,
        "and draws a plot in every one",
      );
    },

    word: (c, is) => {
      is.equal(c.charts.length, LABELS.length, "Word opens the file and builds every chart");
      is.equal(c.chart(0).typeName, "pie", "the first is a pie");
      is.equal(c.chart(1).typeName, "doughnut", "the second a doughnut");
    },

    parity: (p, is) => {
      is.within(p.previewChartWidth(0), p.wordChartWidth(0), "1mm", "the first frame agrees");
      is.within(p.previewChartHeight(0), p.wordChartHeight(0), "1mm", "in both directions");
    },
  },
});
