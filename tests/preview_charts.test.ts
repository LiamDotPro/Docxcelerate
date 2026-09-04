import { test } from "node:test";
import { assertEquals } from "./assert.ts";
import { renderDocxBytes } from "docxcelerate/docx";
import {
  type ChartDrawer,
  type PreviewChart,
  readPackedCharts,
  settleDocxPreviewCharts,
} from "docxcelerate/preview";
import type { DocumentModel, GraphNode } from "docxcelerate";
import { JSDOM } from "jsdom";

/**
 * Reading a chart back out of the file, for the preview to draw.
 *
 * The rule these cases exist to hold is the one the whole preview is built on:
 * **what is shown comes from the packed file, never from the model that
 * produced it.** A chart's colours, its key, the format on its value axis —
 * every one is a decision the packer made from the theme, and a preview that
 * recomputed them would be a second copy of the packer's arithmetic. The copy
 * nobody opens in Word is the one free to drift.
 *
 * So each case packs a document, throws the model away, and asserts on what
 * comes back out of the bytes.
 */

const quarters = ["Q1", "Q2", "Q3", "Q4"];

function documentOf(nodes: GraphNode[]): DocumentModel {
  return {
    schemaVersion: "docxcelerate.letter/v0",
    id: "charts",
    title: "Charts",
    nodes,
  };
}

function chart(over: Partial<GraphNode> = {}): GraphNode {
  return {
    id: "revenue",
    kind: "graph",
    mode: "static",
    graphType: "bar",
    data: {
      categories: quarters,
      series: [
        { label: "2024", values: [12, 18, 9, 22] },
        { label: "2025", values: [16, 14, 21, 27] },
      ],
    },
    ...over,
  } as GraphNode;
}

/** What the file says about the charts in a document, read back out of it. */
async function readBack(nodes: GraphNode[]): Promise<PreviewChart[]> {
  return await readPackedCharts(await renderDocxBytes(documentOf(nodes)));
}

test("a chart comes back out of the file with its numbers intact", async () => {
  const [read] = await readBack([chart()]);

  assertEquals(read.graphType, "bar");
  assertEquals([...read.categories], quarters);
  assertEquals(read.series.length, 2);
  assertEquals(read.series[0].label, "2024");
  assertEquals([...read.series[0].values], [12, 18, 9, 22]);
  assertEquals([...read.series[1].values], [16, 14, 21, 27]);
});

test("a gap in a series is still a gap once it has been through the file", async () => {
  const [read] = await readBack([
    chart({ data: { categories: quarters, series: [{ label: "n", values: [1, null, 3, 4] }] } }),
  ]);

  assertEquals([...read.series[0].values], [1, null, 3, 4]);
});

test("every chart type is recognised on the way back", async () => {
  const types: GraphNode["graphType"][] = [
    "bar",
    "barHorizontal",
    "line",
    "area",
    "pie",
    "doughnut",
    "scatter",
  ];

  for (const graphType of types) {
    const [read] = await readBack([chart({ graphType })]);

    assertEquals(read.graphType, graphType);
  }
});

test("the colours read back are the theme's, in the order it set them", async () => {
  const [read] = await readBack([chart()]);

  assertEquals(read.series[0].color, "2A78D6");
  assertEquals(read.series[1].color, "EB6834");
});

test("a pie's colours are its slices', because that is where the file puts them", async () => {
  const [read] = await readBack([
    chart({
      graphType: "pie",
      data: { categories: quarters, series: [{ label: "Share", values: [4, 3, 2, 1] }] },
    }),
  ]);

  assertEquals(read.sliceColors.length, 4);
  assertEquals(read.sliceColors[0], "2A78D6");
  assertEquals(read.sliceColors[1], "EB6834");
});

test("a chart's own title is read, and an axis's is not mistaken for it", async () => {
  const [titled] = await readBack([
    chart({ title: "Revenue", categoryAxisTitle: "Quarter", valueAxisTitle: "Pounds" }),
  ]);
  const [untitled] = await readBack([chart({ valueAxisTitle: "Pounds" })]);

  assertEquals(titled.title, "Revenue");
  assertEquals(titled.categoryAxisTitle, "Quarter");
  assertEquals(titled.valueAxisTitle, "Pounds");
  // The failure this guards: with no chart title, the first `c:title` in the
  // part is the axis's, and a chart with a labelled axis reported itself as
  // titled after it.
  assertEquals(untitled.title, undefined);
  assertEquals(untitled.valueAxisTitle, "Pounds");
});

test("`General` comes back as no format at all, because that is what it means", async () => {
  const [plain] = await readBack([chart()]);
  const [formatted] = await readBack([chart({ numberFormat: "#,##0" })]);

  // Taken as a format, `General` prints the word in front of every number on
  // the axis — which is exactly what it did before this was mapped.
  assertEquals(plain.numberFormat, undefined);
  assertEquals(formatted.numberFormat, "#,##0");
});

test("the key's place and the stacking come back as the file states them", async () => {
  const [right] = await readBack([chart({ legend: "right", stacked: true })]);
  const [none] = await readBack([chart({ legend: "none" })]);

  assertEquals(right.legend, "right");
  assertEquals(right.stacked, true);
  assertEquals(none.legend, "none");
  assertEquals(none.stacked, false);
});

test("a chart with one series comes back with no key, as the packer decided", async () => {
  const [read] = await readBack([
    chart({ data: { categories: quarters, series: [{ label: "n", values: [1, 2, 3, 4] }] } }),
  ]);

  // The point of the assertion is not the value — it is that the preview takes
  // the packer's decision rather than making the same one again.
  assertEquals(read.legend, "none");
});

test("data labels come back, and which figure they print", async () => {
  const [off] = await readBack([chart()]);
  const [values] = await readBack([chart({ dataLabels: true })]);
  const [percent] = await readBack([
    chart({
      graphType: "pie",
      dataLabels: true,
      data: { categories: quarters, series: [{ label: "Share", values: [4, 3, 2, 1] }] },
    }),
  ]);

  assertEquals(off.dataLabels, "none");
  assertEquals(values.dataLabels, "value");
  assertEquals(percent.dataLabels, "percent");
});

test("the frame comes back in points, which is what the drawer draws for", async () => {
  const [read] = await readBack([chart({ width: 300, height: 150 })]);

  assertEquals(Math.round(read.widthPt), 300);
  assertEquals(Math.round(read.heightPt), 150);
});

test("charts come back in the order their drawings appear", async () => {
  const read = await readBack([
    chart({ id: "one", title: "First" }),
    chart({ id: "two", title: "Second", graphType: "line" }),
    chart({ id: "three", title: "Third", graphType: "pie" }),
  ]);

  assertEquals(read.map((entry) => entry.title), ["First", "Second", "Third"]);
  assertEquals(read.map((entry) => entry.graphType), ["bar", "line", "pie"]);
});

// ---------------------------------------------------------------------------
// Putting the drawing in the frame
// ---------------------------------------------------------------------------

/**
 * A window to build the frames in.
 *
 * The one place these tests need a DOM: what settle does is find elements and
 * put markup in them, and there is no honest way to test that without
 * elements. jsdom is a devDependency and never ships.
 */
let window: JSDOM["window"] | undefined;

function page() {
  window ??= new JSDOM("<!doctype html><html><body></body></html>").window;

  return window;
}

/**
 * The frames docx-preview leaves, written out by hand.
 *
 * A real one is an inline-block span with a width, a height and nothing at all
 * inside it — see `docx_preview_charts`. These cases stand a picture and a
 * shape beside two of them, because the whole reason the frame is matched by
 * shape rather than by a class is that docx-preview gives it none, and matching
 * one too many would put a chart inside somebody's photograph.
 */
function stage(): { body: HTMLElement; frames: HTMLElement[] } {
  const body = page().document.createElement("div");

  body.innerHTML = `
    <p><span><span style="display: inline-block; width: 451pt; height: 263pt;"></span></span></p>
    <p><span><span style="display: inline-block; width: 40pt; height: 40pt;"><img src="x"></span></span></p>
    <p><span><span style=""><svg><rect/></svg></span></span></p>
    <p><span><span style="display: inline-block; width: 200pt; height: 120pt;"></span></span></p>
  `;

  return {
    body,
    // The same rule settle applies, and it has to be: an inline-block span
    // with something already in it is a picture's, not a chart's, and a list
    // that included it would number the frames differently from the code
    // under test.
    frames: [...body.querySelectorAll("span")].filter(
      (span) => span.style.display === "inline-block" && span.childElementCount === 0,
    ) as HTMLElement[],
  };
}

/** A drawer that records what it was asked for and draws something trivial. */
function recorder(markup: (chart: PreviewChart) => string | undefined) {
  const seen: PreviewChart[] = [];
  const draw: ChartDrawer = (chart) => {
    seen.push(chart);
    return markup(chart);
  };

  return { seen, draw };
}

test("each chart is drawn into its own frame, in order", async () => {
  const { body } = stage();
  const charts = await readBack([
    chart({ id: "one", title: "First" }),
    chart({ id: "two", title: "Second" }),
  ]);
  const { seen, draw } = recorder((chart) => `<svg data-title="${chart.title}"></svg>`);

  assertEquals(settleDocxPreviewCharts(body, charts, draw), 2);
  assertEquals(seen.map((entry) => entry.title), ["First", "Second"]);
  assertEquals(
    [...body.querySelectorAll("svg[data-title]")].map((svg) => svg.getAttribute("data-title")),
    ["First", "Second"],
  );
});

test("a picture's frame is not a chart's, however alike they look", async () => {
  const { body } = stage();
  const charts = await readBack([chart({ title: "Only" })]);

  settleDocxPreviewCharts(body, charts, () => `<svg data-drawn="1"></svg>`);

  // The picture keeps its picture: a frame with something already in it is not
  // an empty one, which is the whole of the rule.
  assertEquals(body.querySelectorAll("img").length, 1);
  assertEquals(body.querySelectorAll("svg[data-drawn]").length, 1);
});

test("a drawer that declines leaves the frame as it found it", async () => {
  const { body, frames } = stage();
  const charts = await readBack([chart(), chart({ id: "two" })]);

  assertEquals(settleDocxPreviewCharts(body, charts, () => undefined), 0);
  assertEquals(frames[0].childElementCount, 0);
  // An empty frame of the right size is a gap. A wrong plot would be a lie,
  // and the frame is what the page is laid out around either way.
  assertEquals(frames[0].style.width, "451pt");
});

test("a chart says what it is for a reader being read the page", async () => {
  const { body, frames } = stage();
  const charts = await readBack([chart({ title: "Revenue by quarter" })]);

  settleDocxPreviewCharts(body, charts, () => "<svg></svg>");

  assertEquals(frames[0].getAttribute("role"), "img");
  assertEquals(frames[0].getAttribute("aria-label"), "Revenue by quarter: 2024, 2025");
});

test("more frames than charts leaves the spare frames alone", async () => {
  const { body, frames } = stage();
  const charts = await readBack([chart()]);

  assertEquals(settleDocxPreviewCharts(body, charts, () => "<svg></svg>"), 1);
  assertEquals(frames[1].childElementCount, 0);
});
