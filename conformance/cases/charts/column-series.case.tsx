import { Document, Graph, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle, COLUMN_MM } from "../_support/style.ts";

/**
 * A chart, packed as a chart.
 *
 * This is the case the whole chart feature stands on, and what it is really
 * asserting is a negative: that the document does **not** hold a picture of a
 * chart. Every other way of getting a plot into a `.docx` — rendering to PNG,
 * to SVG, to a group of drawn shapes — produces something a reader can look at
 * and nothing else. A `c:chartSpace` produces a chart Word built: selectable,
 * restyleable, re-typeable, printed as vectors at whatever resolution the
 * printer has, and carrying its own numbers so "Edit Data" opens them.
 *
 * A chart is also the only feature in the suite that is not in
 * `document.xml`. It is four parts that have to agree — the drawing, the chart
 * part it names, the relationship joining them and the content-type override
 * declaring it — plus a workbook hanging off the chart. Any one of them
 * missing gives a file that either draws nothing or that Word offers to
 * repair, and three of the four still look right when the fourth is gone. So
 * probe A follows the whole chain rather than asserting on the drawing.
 *
 * **The values are asserted from the cache, not from the workbook.** The cache
 * is what Word draws from without opening anything; a chart whose numbers live
 * only in the workbook draws as an empty frame until Excel is asked for them.
 * The workbook is asserted to exist, because it is what makes the chart
 * editable rather than what makes it draw.
 *
 * **The preview draws the frame from the file and the plot from a library.**
 * docx-preview has no reading of a chart part at all — the string "chart" does
 * not occur in its bundle — so what it leaves is an empty inline-block span at
 * exactly the size the file gave the frame. The frame is the half that has to
 * be exact, because it is what the page is laid out around and what keeps the
 * preview paginating like Word; the plot inside it is drawn by ECharts from
 * the same packed bytes, through the very drawer a scaffolded workspace uses.
 *
 * That is why the parity tier here compares frames and not pixels. Holding
 * ECharts' bars against Word's would be measuring two renderers rather than
 * measuring the document, and it would fail for reasons no document could fix.
 * What parity says is the thing that matters: the space the preview lays the
 * page out around is the space Word lays it out around.
 */
const QUARTERS = ["Q1", "Q2", "Q3", "Q4"];
const LAST_YEAR = [12400, 18100, 9300, 22800];
const THIS_YEAR = [16200, 14500, 21700, 27300];

/** The first two colours of the shipped chart palette, in order. */
const SERIES_ONE = "2A78D6";
const SERIES_TWO = "EB6834";

const HEIGHT_PT = 200;

export default defineCase({
  id: "charts/column-series",
  feature: "chart.columnSeries",
  title: "A clustered column chart of two series, drawn by Word from its own data",
  word: "Insert → Chart → Clustered Column (c:chartSpace / c:barChart)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="column-series" title="Charts">
      <Paragraph id="before">A paragraph above the chart.</Paragraph>

      <Graph
        id="revenue"
        title="Revenue by quarter"
        graphType="bar"
        height={HEIGHT_PT}
        numberFormat="#,##0"
        categoryAxisTitle="Quarter"
        valueAxisTitle="Revenue"
        caption="Figure 1: revenue by quarter."
        data={{
          categories: QUARTERS,
          series: [
            { label: "2024", values: LAST_YEAR },
            { label: "2025", values: THIS_YEAR },
          ],
        }}
      />

      <Paragraph id="after">A paragraph below it.</Paragraph>
    </Document>
  ),

  regions: [{ id: "revenue", anchor: "Figure 1" }],

  expect: {
    ooxml: (a, is) => {
      const xml = a.documentXml ?? "";
      const chart = a.chart(0);

      // The chain, part by part. Each of these is a separate way for a chart
      // to be in the model and not in the file.
      is.includes(xml, "<w:drawing>", "the body holds a drawing");
      is.includes(
        xml,
        'uri="http://schemas.openxmlformats.org/drawingml/2006/chart"',
        "declaring itself a chart rather than a picture",
      );
      is.excludes(xml, "dxclChart_", "and its relationship token became a real id");
      is.equal(chart.partPresent, true, "the chart part the drawing names is in the package");
      is.equal(chart.workbookPresent, true, "and so is the workbook Edit Data opens");
      is.equal(chart.externalData, true, "the chart part names that workbook");

      // What Word will draw.
      is.equal(chart.plot, "bar", "the plot is a bar chart");
      is.equal(chart.barDir, "col", "with the bars standing up");
      is.equal(chart.grouping, "clustered", "clustered rather than stacked");
      is.equal(chart.axisCount, 2, "and two axes");

      // The numbers. Cached, because the cache is what Word draws from.
      is.equal(chart.categories, QUARTERS, "every category is cached in the file");
      is.equal(a.chartSeries(0, 0).label, "2024", "the first series is named");
      is.equal(a.chartSeries(0, 0).values, LAST_YEAR, "and every one of its values is cached");
      is.equal(a.chartSeries(0, 1).label, "2025", "and the second");
      is.equal(a.chartSeries(0, 1).values, THIS_YEAR, "with its values too");

      // How it is themed. The colours come from the palette in order, and the
      // ink from the theme — a chart that lettered itself in the series'
      // colours would be one nobody could re-theme.
      is.equal(a.chartSeries(0, 0).color, SERIES_ONE, "the first series takes the palette's first colour");
      is.equal(a.chartSeries(0, 1).color, SERIES_TWO, "and the second its second");
      is.includes(chart.xml ?? "", `<a:srgbClr val="${caseStyle.palette?.rule}"/>`, "the grid is ruled in the theme's rule colour");
      is.includes(chart.xml ?? "", `<a:srgbClr val="${caseStyle.palette?.muted}"/>`, "and lettered in its muted ink");
      is.includes(chart.xml ?? "", `<a:latin typeface="${caseStyle.typography.bodyFont}"/>`, "in the document's own face");

      // What the document said about it.
      is.equal(chart.title, "Revenue by quarter", "the chart carries its heading");
      is.equal(chart.autoTitleDeleted, "0", "so Word is told to draw one");
      is.equal(chart.legend, "b", "the key sits under the plot");
      is.equal(chart.numberFormat, "#,##0", "the value axis prints the format the document asked for");
      is.includes(chart.xml ?? "", "<a:t>Quarter</a:t>", "the category axis is labelled");
      is.includes(chart.xml ?? "", "<a:t>Revenue</a:t>", "and the value axis");
      is.equal(chart.description, "Revenue by quarter", "and the drawing says what it is for a reader who cannot see it");

      // The frame, and the caption beside rather than inside it.
      is.within(chart.widthPt, COLUMN_MM * (72 / 25.4), 1, "the chart fills the text column");
      is.within(chart.heightPt, HEIGHT_PT, 1, "at the depth the document gave it");
      is.includes(xml, "Figure 1: revenue by quarter.", "the caption is a paragraph in the body");
      is.excludes(chart.xml ?? "", "Figure 1", "and not text drawn into the chart");
    },

    preview: (b, is) => {
      const chart = b.chart(0);

      // The frame is the half the preview gets right, and it is the half
      // pagination depends on. A chart frame of the wrong size lays the page
      // out differently from Word whether or not a plot is drawn in it.
      is.equal(b.charts.length, 1, "the preview lays out one chart frame");
      is.within(chart.w, b.mm(COLUMN_MM), "1mm", "as wide as the text column");
      is.within(chart.h, b.pt(HEIGHT_PT), "1mm", "and as deep as the file says");
      is.equal(chart.picture, false, "and not as a picture");

      // And a plot is actually drawn in it. Without the drawer this is where
      // the preview shows a hole, and a hole looks deliberate.
      is.equal(chart.plotted, true, "and draws the chart inside it");
    },

    word: (c, is) => {
      const chart = c.chart(0);

      // The tier that settles it. Everything above says the file describes a
      // chart; only this says Word built one.
      is.equal(c.charts.length, 1, "Word finds one chart");
      is.equal(chart.typeName, "columnClustered", "and builds a clustered column from it");
      is.equal(chart.seriesCount, 2, "with two series");
      is.equal(chart.title, "Revenue by quarter", "titled as the document said");
      is.equal(chart.hasLegend, true, "and keyed");

      // The assertion nothing else in the suite can make: Word reading the
      // document's own numbers back out of the chart it built.
      is.equal(c.chartSeries(0, 0).name, "2024", "Word reads the first series by name");
      is.equal(c.chartSeries(0, 0).values, LAST_YEAR, "and every figure in it");
      is.equal(c.chartSeries(0, 1).name, "2025", "and the second");
      is.equal(c.chartSeries(0, 1).values, THIS_YEAR, "with its figures too");
      is.equal(chart.categories, QUARTERS, "counted against the document's own categories");

      is.equal(c.chartSeries(0, 0).color, SERIES_ONE, "drawn in the palette's first colour");
      is.within(chart.width, c.mm(COLUMN_MM), 1, "the frame fills the text column");
      is.within(chart.height, HEIGHT_PT, 1, "at the depth the document gave it");
      // A plot area of nothing is what a chart Word could not lay out reports,
      // and it looks identical to a correct one in every other measurement.
      is.greater(chart.plotWidth, 0, "and Word lays out a plot inside it");
      is.greater(chart.plotHeight, 0, "with a height");
    },

    /**
     * The frame on screen against the frame in Word.
     *
     * Only the frame, and permanently. Both sides draw a plot inside it and
     * neither drew the other's: Word's is Word's, the preview's is ECharts'.
     * Holding one against the other pixel for pixel would be a measurement of
     * two chart libraries, red for reasons no document could fix. The frame is
     * the fact both sides take from the file, and it is the one the page is
     * laid out around — so it is the one that has to agree.
     */
    parity: (p, is) => {
      is.within(p.previewChartWidth(0), p.wordChartWidth(0), "1mm", "the frame is as wide on screen as in Word");
      is.within(p.previewChartHeight(0), p.wordChartHeight(0), "1mm", "and as deep");
    },
  },
});
