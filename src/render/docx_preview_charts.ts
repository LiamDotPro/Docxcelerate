/**
 * Drawing a document's charts into the preview, from the packed file.
 *
 * docx-preview has no reading of a chart part at all — the string "chart" does
 * not occur in its bundle — so what it leaves where a chart goes is an empty
 * inline-block span at exactly the size the file gave the frame. The frame
 * being right is what keeps the preview paginating like Word. The hole inside
 * it is what this fills.
 *
 * **The chart is read out of the package, never out of the model.** That is the
 * same rule {@linkcode ../render/docx_packed.ts | docx_packed} follows and for
 * the same reason: the colours a series is drawn in, the key a chart with one
 * series does not get, the format the value axis prints — every one of those is
 * a decision the packer made from the theme, and recomputing them here would be
 * a second copy of the packer's arithmetic living beside the first. The copy
 * nobody opens in Word is the one free to drift.
 *
 * **The drawing itself is somebody else's.** This module reads and it places;
 * it does not plot. A chart renderer is a large and well-solved problem, and
 * bundling one into a package whose only dependency is its packer would be the
 * wrong trade for every caller who never previews a chart. So the caller passes
 * a {@linkcode ChartDrawer} — a scaffolded workspace wires one up over the
 * charting library it already has in the browser — and this puts what comes
 * back in the frame the file declared.
 *
 * What the preview then shows is honest about what it is: the frame, the data,
 * the colours and the type are the file's, and the plot inside is another
 * renderer's drawing of them rather than Word's. Word is the only thing that
 * draws what Word draws.
 *
 * @module
 */

import type { GraphLegend, GraphType } from "../domain/types.ts";
import { readPart } from "./docx_packed.ts";
import { attribute, decodeEntities, element, elements, numberAttribute } from "./ooxml_read.ts";

/** One plotted run of numbers, as the packed chart part records it. */
export interface PreviewChartSeries {
  /** What the key calls it, where it has a name. */
  readonly label?: string;
  /** Its numbers, with a reading nobody took kept as `null`. */
  readonly values: readonly (number | null)[];
  /** The colour the file draws it in, as RGB hex without the `#`. */
  readonly color?: string;
}

/** One chart, as the packed file describes it. */
export interface PreviewChart {
  /** Which plot Word will draw. */
  readonly graphType: GraphType;
  /** The chart's own heading, where it has one. */
  readonly title?: string;
  /** Where the key sits, or `"none"` where the file declares none. */
  readonly legend: GraphLegend;
  /** Whether the series stack rather than stand beside one another. */
  readonly stacked: boolean;
  /**
   * The OOXML format the value axis prints its numbers in.
   *
   * Absent where the chart states none. `General` is not a format — it is
   * OOXML's way of writing "no format", and a drawer handed it as one prints
   * the word in front of every number on the axis.
   */
  readonly numberFormat?: string;
  /** Whether each point prints its own figure, and which figure. */
  readonly dataLabels: "none" | "value" | "percent";
  /** What the category axis counts along. */
  readonly categoryAxisTitle?: string;
  /** What the value axis measures. */
  readonly valueAxisTitle?: string;
  /** The categories the values are counted against. */
  readonly categories: readonly string[];
  /** The series, in the order they are drawn and keyed. */
  readonly series: readonly PreviewChartSeries[];
  /**
   * The colours a pie's slices are painted in, in order.
   *
   * A pie has one series and many colours, so its palette is on the data
   * points rather than on the series. Empty for every other kind.
   */
  readonly sliceColors: readonly string[];
  /** How wide the frame is, in points. */
  readonly widthPt: number;
  /** How deep it is, in points. */
  readonly heightPt: number;
  /** The ink the chart's own text is set in. */
  readonly textColor?: string;
  /** The colour its gridlines and axes are ruled in. */
  readonly ruleColor?: string;
  /** The face its text is set in. */
  readonly font?: string;
}

/**
 * What draws one chart.
 *
 * Handed a chart read out of the file and expected to return SVG or HTML for
 * the frame it goes in — sized to `widthPt` by `heightPt`, because that is the
 * space the page is already laid out around. Returning `undefined` leaves the
 * frame as docx-preview left it, which is the right answer for a chart the
 * drawer does not know how to plot: an empty frame of the right size is a gap,
 * and a wrong plot is a lie.
 */
export type ChartDrawer = (chart: PreviewChart) => string | undefined;

/** The relationship type a drawing uses to reach a chart part. */
const CHART_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

/** A point, in the English Metric Units a drawing states its size in. */
const EMU_PER_PT = 12700;

/**
 * Every chart the body draws, in document order.
 *
 * The running strips are left out on purpose: docx-preview renders them into
 * their own containers, and the frames this matches against are the body's.
 *
 * @param packed The `.docx` bytes — the same ones handed to `renderAsync`.
 * @returns One entry per chart, in the order the drawings appear.
 *
 * @example
 * ```ts
 * const bytes = new Uint8Array(await blob.arrayBuffer());
 * await renderAsync(bytes, body, head, { breakPages: true });
 * settleDocxPreviewCharts(body, await readPackedCharts(bytes), drawWithEcharts);
 * ```
 */
export async function readPackedCharts(packed: Uint8Array): Promise<PreviewChart[]> {
  const document = await readPart(packed, "word/document.xml");

  if (document === undefined) {
    return [];
  }

  const targets = await relationshipsOf(packed, "word/_rels/document.xml.rels");
  const found: PreviewChart[] = [];

  for (const drawing of elements(document, "w:drawing")) {
    const reference = element(drawing, "c:chart");
    const id = reference === null ? undefined : attribute(reference, "r:id");
    const target = id === undefined ? undefined : targets.get(id);

    if (target === undefined) {
      continue;
    }

    const xml = await readPart(packed, `word/${target.replace(/^\.\//, "")}`);

    if (xml === undefined) {
      continue;
    }

    const extent = element(drawing, "wp:extent") ?? "";

    found.push({
      ...readChart(xml),
      widthPt: (numberAttribute(extent, "cx") ?? 0) / EMU_PER_PT,
      heightPt: (numberAttribute(extent, "cy") ?? 0) / EMU_PER_PT,
    });
  }

  return found;
}

/**
 * Draws a document's charts into the frames docx-preview left for them.
 *
 * Call it after `renderAsync` and before the container is measured,
 * screenshotted or paginated — the same point {@linkcode settleDocxPreview} is
 * called at, and for the same reason: a chart that arrives after the page is
 * measured is a page measured around a hole.
 *
 * @param container The element `renderAsync` rendered into.
 * @param charts What the file says about its charts, from
 * {@linkcode readPackedCharts}.
 * @param draw What plots one. See {@linkcode ChartDrawer}.
 * @returns How many frames were filled, which is fewer than `charts.length`
 * when the drawer declined one.
 */
export function settleDocxPreviewCharts(
  container: Element,
  charts: readonly PreviewChart[],
  draw: ChartDrawer,
): number {
  const frames = chartFrames(container);
  let drawn = 0;

  for (const [index, frame] of frames.entries()) {
    const chart = charts[index];

    if (chart === undefined) {
      continue;
    }

    const markup = draw(chart);

    if (markup === undefined || markup === "") {
      continue;
    }

    frame.innerHTML = markup;
    // What the frame holds is a picture of data, and its own labels are drawn
    // inside it as shapes rather than as text a reader can be handed. So the
    // frame says what it is, once, the way the drawing in the file already
    // does — and the plot inside is hidden from anything reading the page out.
    frame.setAttribute("role", "img");
    frame.setAttribute("aria-label", describe(chart));
    drawn += 1;
  }

  return drawn;
}

/**
 * The empty frames docx-preview leaves where the charts go.
 *
 * Matched by shape rather than by a class, because docx-preview gives them
 * none: a chart comes out as an inline-block span with a width and a height
 * and nothing at all inside it. A picture is the same span with an `<img>` in
 * it and a shape is an `<svg>` with no inline-block, so neither is caught by
 * this — which is what makes the nth frame here the nth chart in the file.
 */
function chartFrames(container: Element): HTMLElement[] {
  // Duck-typed rather than `instanceof HTMLElement`. A preview is built in one
  // document and moved into an iframe to be paginated, and an element carried
  // across realms is not an instance of the constructor this realm has — the
  // check would quietly match nothing exactly where the preview is being put
  // together.
  return [...container.querySelectorAll("span")]
    .filter((span): span is HTMLElement => {
      const style = (span as HTMLElement).style as CSSStyleDeclaration | undefined;

      return style !== undefined &&
        style.display === "inline-block" &&
        style.width !== "" &&
        style.height !== "" &&
        span.childElementCount === 0;
    });
}

/** What a chart is called for a reader who is being read the page. */
function describe(chart: PreviewChart): string {
  const named = chart.series
    .map((series) => series.label)
    .filter((label): label is string => label !== undefined);

  const what = chart.title ?? `${chart.graphType} chart`;

  return named.length === 0 ? what : `${what}: ${named.join(", ")}`;
}

/** A relationships part as a map of id to target. */
async function relationshipsOf(
  packed: Uint8Array,
  part: string,
): Promise<Map<string, string>> {
  const xml = await readPart(packed, part);
  const found = new Map<string, string>();

  if (xml === undefined) {
    return found;
  }

  for (const entry of elements(xml, "Relationship")) {
    const id = attribute(entry, "Id");
    const target = attribute(entry, "Target");

    if (id !== undefined && target !== undefined && entry.includes(CHART_RELATIONSHIP)) {
      found.set(id, target);
    }
  }

  return found;
}

// ---------------------------------------------------------------------------
// Reading one chart part
//
// Only the parts of it a drawing needs. A chart part carries a great deal that
// exists to tell Word how to draw — text properties per axis, effect lists,
// layout hints — and none of it belongs in a shape another renderer is handed.
// What comes out is what the chart *is*: the plot, the numbers, and the colours
// the theme chose.
// ---------------------------------------------------------------------------

/** Which plot a chart part declares, mapped back to the name the model uses. */
function graphTypeOf(xml: string): GraphType {
  if (element(xml, "c:doughnutChart") !== null) return "doughnut";
  if (element(xml, "c:pieChart") !== null) return "pie";
  if (element(xml, "c:lineChart") !== null) return "line";
  if (element(xml, "c:areaChart") !== null) return "area";
  if (element(xml, "c:scatterChart") !== null) return "scatter";

  return attribute(element(xml, "c:barDir") ?? "", "val") === "bar" ? "barHorizontal" : "bar";
}

function readChart(xml: string): Omit<PreviewChart, "widthPt" | "heightPt"> {
  const graphType = graphTypeOf(xml);
  const pie = graphType === "pie" || graphType === "doughnut";
  // An axis carries a `c:title` of its own, so the chart's is the one standing
  // before the plot area. Taking the first in the part would report a chart
  // with no heading and a labelled axis as being titled after the axis.
  const head = xml.slice(0, xml.indexOf("<c:plotArea>"));
  // The value axis is written last, which on a scatter — the one chart with
  // two of them — is what separates it from the x axis. The axis the
  // categories run along is whichever comes first, and a pie has neither.
  const valueAxes = elements(xml, "c:valAx");
  const valueAxis = valueAxes.at(-1) ?? "";
  const categoryAxis = elements(xml, "c:catAx")[0] ?? (valueAxes.length === 2 ? valueAxes[0] : "");
  const series = elements(xml, "c:ser");
  const first = series[0] ?? "";

  return {
    graphType,
    title: titleOf(head),
    legend: legendOf(attribute(element(xml, "c:legendPos") ?? "", "val")),
    stacked: attribute(element(xml, "c:grouping") ?? "", "val") === "stacked",
    numberFormat: formatOf(attribute(element(valueAxis, "c:numFmt") ?? "", "formatCode")),
    dataLabels: labelsOf(element(xml, "c:dLbls") ?? ""),
    categoryAxisTitle: titleOf(categoryAxis),
    valueAxisTitle: titleOf(valueAxis),
    // A scatter counts along numbers rather than names, so its categories are
    // its x values written out — which is what a drawer wants to put under the
    // points either way.
    categories: graphType === "scatter"
      ? cachedNumbers(element(first, "c:xVal") ?? "").map((value) => String(value ?? ""))
      : cachedText(element(first, "c:cat") ?? ""),
    series: series.map((entry) => ({
      label: cachedText(element(entry, "c:tx") ?? "")[0],
      values: cachedNumbers(element(entry, "c:val") ?? element(entry, "c:yVal") ?? ""),
      // The series' own paint, not a data point's: on a pie the two differ,
      // and `c:spPr` is where the series states the one colour it has.
      color: colorOf(element(entry, "c:spPr") ?? ""),
    })),
    // The slice colours, which a pie keeps on its data points rather than on
    // its one series.
    sliceColors: pie
      ? elements(first, "c:dPt")
        .map((point) => colorOf(element(point, "c:spPr") ?? ""))
        .filter((color): color is string => color !== undefined)
      : [],
    textColor: colorOf(element(xml, "c:txPr") ?? ""),
    ruleColor: colorOf(element(xml, "c:majorGridlines") ?? ""),
    font: attribute(element(xml, "a:latin") ?? "", "typeface"),
  };
}

/**
 * A number format, or nothing where the chart states none.
 *
 * `General` is what OOXML writes for "however this number looks by itself", so
 * it is a statement that there is no format rather than a format — and a
 * drawer handed it as one prints the word "General" in front of every number
 * on the axis, which is exactly what happened before this existed.
 */
function formatOf(format: string | undefined): string | undefined {
  return format === undefined || format === "General" ? undefined : format;
}

/** Whether each point prints its own figure, and which figure it prints. */
function labelsOf(labels: string): "none" | "value" | "percent" {
  const on = (name: string) => attribute(element(labels, name) ?? "", "val") === "1";

  if (on("c:showPercent")) return "percent";

  return on("c:showVal") ? "value" : "none";
}

/** Where a key sits, as the model names it, or that the file declares none. */
function legendOf(position: string | undefined): GraphLegend {
  const named: Record<string, GraphLegend> = { t: "top", b: "bottom", l: "left", r: "right" };

  return position === undefined ? "none" : named[position] ?? "bottom";
}

/** The words a `c:title` prints, joined as one line. */
function titleOf(xml: string): string | undefined {
  const title = element(xml, "c:title");

  if (title === null) {
    return undefined;
  }

  const runs = [...title.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((match) =>
    decodeEntities(match[1])
  );

  return runs.length === 0 ? undefined : runs.join("");
}

/** The first colour stated in a run of XML. */
function colorOf(xml: string): string | undefined {
  return /<a:srgbClr val="([0-9A-Fa-f]{6})"\s*\/>/.exec(xml)?.[1];
}

/** The strings a reference caches, in index order. */
function cachedText(reference: string): string[] {
  return elements(reference, "c:pt").map((point) =>
    decodeEntities((element(point, "c:v") ?? "").replace(/<[^>]*>/g, ""))
  );
}

/**
 * The numbers a reference caches, with a missing point kept as a gap.
 *
 * The count says how many readings there are; a point that is absent is one
 * nobody took, and reading it as a zero would draw a bar for a month that has
 * not happened.
 */
function cachedNumbers(reference: string): (number | null)[] {
  const cache = element(reference, "c:numCache") ?? reference;
  const count = numberAttribute(element(cache, "c:ptCount") ?? "", "val") ?? 0;
  const values: (number | null)[] = new Array(count).fill(null);

  for (const point of elements(cache, "c:pt")) {
    const at = numberAttribute(point, "idx");
    const raw = (element(point, "c:v") ?? "").replace(/<[^>]*>/g, "");
    const value = Number.parseFloat(raw);

    if (at !== undefined && Number.isFinite(value)) {
      values[at] = value;
    }
  }

  return values;
}
