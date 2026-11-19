import { Document, Graph, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * One bar and one slice picked out in a colour of their own.
 *
 * A series colour paints the whole series; the only way to single out one
 * point is a `c:dPt` inside the series, which is also what keeps the point
 * selectable and restyleable in Word. The XML follows the schema's element
 * order, and a `c:dPt` out of order or with a flag Word dislikes is a file it
 * offers to repair, so the Word tier is the proof that it opens. The Word probe
 * does not read point colours, so they are asserted from the file only.
 */
const RED = "C0392B";

export default defineCase({
  id: "charts/point-colors",
  feature: "chart.pointColors",
  title: "A bar and a pie slice picked out in their own colour",
  word: "Format Data Point → Fill (c:ser/c:dPt)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="point-colors" title="Charts">
      <Graph
        id="ageing"
        title="Receivables by age"
        graphType="bar"
        height={200}
        caption="Figure 1: receivables by age."
        data={{
          categories: ["Current", "30 days", "60 days", "Overdue"],
          series: [{ label: "Owed", values: [18200, 6400, 2100, 3900], pointColors: [null, null, null, RED] }],
        }}
      />
      <Graph
        id="status"
        title="Invoices by status"
        graphType="pie"
        height={200}
        caption="Figure 2: invoices by status."
        data={{
          categories: ["Paid", "Due", "Overdue"],
          series: [{ values: [62, 25, 13], pointColors: [null, null, RED] }],
        }}
      />
      <Paragraph id="after">A paragraph below them.</Paragraph>
    </Document>
  ),

  regions: [{ id: "ageing", anchor: "Figure 1" }, { id: "status", anchor: "Figure 2" }],

  expect: {
    ooxml: (a, is) => {
      is.includes(
        a.chart(0).xml ?? "",
        `<c:dPt><c:idx val="3"/><c:invertIfNegative val="0"/><c:spPr><a:solidFill><a:srgbClr val="${RED}"/>`,
        "the overdue bar is its own data point",
      );
      is.equal(a.chartSeries(0, 0).color, "2A78D6", "the rest of the series keeps the palette");
      is.includes(
        a.chart(1).xml ?? "",
        `<c:dPt><c:idx val="2"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${RED}"/>`,
        "the overdue slice is painted red",
      );
      is.includes(
        a.chart(1).xml ?? "",
        '<c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="2A78D6"/>',
        "and the slice beside it keeps the palette",
      );
    },

    preview: (b, is) => {
      is.equal(b.charts.length, 2, "the preview lays out both chart frames");
      is.equal(b.chart(0).plotted, true, "and draws the bar chart");
      is.equal(b.chart(1).plotted, true, "and the pie");
    },

    word: (c, is) => {
      is.equal(c.charts.length, 2, "Word opens the file and finds both charts");
      is.equal(c.chart(0).typeName, "columnClustered", "a clustered column");
      is.equal(c.chart(1).typeName, "pie", "and a pie");
    },
  },
});
