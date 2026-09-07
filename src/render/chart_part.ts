/**
 * A chart node, written out as the DrawingML part Word draws from.
 *
 * A chart is packed as a chart, not as a picture of one. The numbers travel in
 * the file, so the reader can select it, restyle it, change its type and open
 * its data — and it is drawn by Word at whatever size and resolution the page
 * and the printer need, rather than at whichever one a build guessed. Nothing
 * here rasterises anything, which is also why the toolkit needs no drawing
 * library and no native module to produce a chart.
 *
 * Two things come out of a chart node. The chart part itself holds the plot's
 * shape and a cache of every value, which is what Word draws from without
 * opening anything. The workbook beside it holds the same numbers as a
 * spreadsheet, which is what "Edit Data" opens. Word needs both: the cache
 * alone renders, but a chart whose data cannot be opened is a picture with
 * extra steps.
 *
 * @module
 */

import type {
  DocumentStyle,
  GraphData,
  GraphLegend,
  GraphNode,
  GraphSeries,
} from "../domain/types.ts";
import type { ZipEntry } from "./ooxml_zip.ts";

/**
 * The colours a chart draws its series in when the theme names none.
 *
 * Eight hues in a fixed order, stepped so that neighbouring pairs stay apart
 * for a colourblind reader as well as a sighted one — worst adjacent pair 9.1
 * ΔE under protanopia, 19.6 in normal vision. The order is the safety, not a
 * preference: rotating it or generating a ninth hue undoes the property the
 * set was chosen for. A chart needing more than eight series wants fewer
 * series.
 *
 * Three of them sit under 3:1 against white, which is why every chart drawn
 * here keeps its axis labels and its key: identity is never left to colour
 * alone.
 */
export const DEFAULT_SERIES_COLORS: readonly string[] = [
  "2A78D6",
  "EB6834",
  "1BAF7A",
  "EDA100",
  "E87BA4",
  "008300",
  "4A3AA7",
  "E34948",
];

/** The ink a chart's own text is set in when the theme names none. */
const DEFAULT_TEXT_COLOR = "595959";
/** The colour a chart rules its gridlines and axes in when the theme names none. */
const DEFAULT_RULE_COLOR = "D9D9D9";
/** How heavy a gridline is drawn, in EMU. Three quarters of a point. */
const GRIDLINE_WIDTH_EMU = 9525;
/** The size a chart's axis and key text is set at, in hundredths of a point. */
const CHART_TEXT_SIZE = 900;
/** The size a chart's title is set at, in hundredths of a point. */
const CHART_TITLE_SIZE = 1200;
/** The id of the category axis, referenced by the plot and by the value axis. */
const CATEGORY_AXIS_ID = 111111111;
/** The id of the value axis. */
const VALUE_AXIS_ID = 222222222;
/** The sheet a chart's formulas name, which is the one the workbook holds. */
const SHEET = "Sheet1";

/** A chart node written out: the part Word draws from, and the data behind it. */
export interface ChartPart {
  /** The `c:chartSpace` XML — the chart's shape and its cached values. */
  readonly chart: string;
  /**
   * The workbook the chart's data opens from, as its own package's parts.
   *
   * Returned unzipped so the caller writes one zip inside another with the
   * same writer, rather than this module owning a second copy of it.
   */
  readonly workbook: readonly ZipEntry[];
}

/**
 * Whether a chart node has anything to plot.
 *
 * A dynamic chart the engine has not filled yet has no series, and one that
 * came back empty is the same thing. Either way there is no chart to draw, and
 * the node falls back to its placeholder rather than packing an empty frame.
 *
 * @param node The chart node.
 * @returns Whether it holds at least one series with at least one value.
 */
export function hasChartData(node: GraphNode): boolean {
  const series = node.data?.series;

  return Array.isArray(series) && series.some((entry) => (entry?.values?.length ?? 0) > 0);
}

/**
 * Writes a chart node out as the parts a Word package needs for it.
 *
 * @param node The chart node, which must have data — see
 * {@linkcode hasChartData}.
 * @param style The document's style, which the chart takes its colours and its
 * type from so that re-theming a document re-themes its charts.
 * @returns The chart's XML and the workbook behind it.
 */
export function chartPartOf(node: GraphNode, style: DocumentStyle): ChartPart {
  const data = normalised(node.data);
  const colors = seriesColors(style);
  const palette = {
    text: style.palette?.muted ?? DEFAULT_TEXT_COLOR,
    heading: style.palette?.heading ?? DEFAULT_TEXT_COLOR,
    rule: style.palette?.rule ?? DEFAULT_RULE_COLOR,
    font: style.typography.bodyFont,
  };

  return {
    chart: chartSpaceXml(node, data, colors, palette),
    workbook: workbookOf(node.graphType, data),
  };
}

/** The colours a document draws series in: the theme's, or the shipped set. */
function seriesColors(style: DocumentStyle): readonly string[] {
  const named = style.palette?.series;

  return named !== undefined && named.length > 0 ? named : DEFAULT_SERIES_COLORS;
}

/** How a chart's text is set, once, so every part of it agrees. */
interface ChartPalette {
  readonly text: string;
  readonly heading: string;
  readonly rule: string;
  readonly font: string;
}

/**
 * A chart's data with every series the same length as the categories.
 *
 * A series shorter than the categories is a series with gaps at the end, and a
 * series longer than them is one whose extra readings have nothing to be
 * counted against. Squaring it here means everything downstream — the caches,
 * the formulas, the workbook's rows — is one rectangle, and none of them has
 * to decide separately what a ragged edge means.
 */
function normalised(data: GraphData | undefined): {
  categories: string[];
  series: GraphSeries[];
} {
  const series = (data?.series ?? []).filter((entry): entry is GraphSeries => entry != null);
  const longest = series.reduce((most, entry) => Math.max(most, entry.values.length), 0);
  const given = data?.categories ?? [];
  const categories = Array.from(
    { length: Math.max(longest, given.length) },
    (_, index) => given[index] ?? String(index + 1),
  );

  return {
    categories,
    series: series.map((entry) => ({
      ...entry,
      values: categories.map((_, index) => entry.values[index] ?? null),
      // Squared up the same way, and only where the series has any: a series
      // with no sizes is every chart but a bubble, and giving it a row of
      // nulls would write an empty `c:bubbleSize` into charts that have no
      // such element.
      ...(entry.sizes === undefined
        ? {}
        : { sizes: categories.map((_, index) => entry.sizes?.[index] ?? null) }),
    })),
  };
}

// ---------------------------------------------------------------------------
// The chart part
// ---------------------------------------------------------------------------

/** The whole `c:chartSpace`, in the order the schema requires it. */
function chartSpaceXml(
  node: GraphNode,
  data: { categories: string[]; series: GraphSeries[] },
  colors: readonly string[],
  palette: ChartPalette,
): string {
  const parts = [
    `<c:date1904 val="0"/>`,
    `<c:roundedCorners val="0"/>`,
    `<c:chart>`,
    titleXml(node.title, palette),
    `<c:autoTitleDeleted val="${node.title === undefined ? 1 : 0}"/>`,
    `<c:plotArea><c:layout/>`,
    plotXml(node, data, colors),
    axesXml(node, palette),
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>`,
    `</c:plotArea>`,
    legendXml(legendOf(node, data.series.length), palette),
    `<c:plotVisOnly val="1"/>`,
    // A missing reading is a gap in the line, not a drop to zero. `zero` is
    // what draws a month nobody measured as a month that measured nothing.
    `<c:dispBlanksAs val="gap"/>`,
    `</c:chart>`,
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>`,
    textPropertiesXml(palette, CHART_TEXT_SIZE),
    // The workbook is related from the chart's own rels, always under this id,
    // because the chart part has exactly one relationship of its own.
    `<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>`,
  ];

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"` +
    ` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
    ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    parts.join("") +
    `</c:chartSpace>`;
}

/** The chart's heading, which is the node's title. */
function titleXml(title: string | undefined, palette: ChartPalette): string {
  if (title === undefined) {
    return "";
  }

  return `<c:title><c:tx><c:rich><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${CHART_TITLE_SIZE}" b="0"><a:solidFill><a:srgbClr val="${palette.heading}"/></a:solidFill><a:latin typeface="${
    escapeXml(palette.font)
  }"/></a:defRPr></a:pPr><a:r><a:t>${
    escapeXml(title)
  }</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:title>`;
}

/** Where the key sits: what the node said, or what the chart needs. */
function legendOf(node: GraphNode, count: number): GraphLegend {
  if (node.legend !== undefined) {
    return node.legend;
  }

  // A pie's key names its slices, so it keeps one however few there are. Every
  // other chart with a single series has a key that repeats its own title.
  return node.graphType === "pie" || node.graphType === "doughnut" || count > 1
    ? "bottom"
    : "none";
}

/** The key, drawn in the document's own ink rather than the series' colours. */
function legendXml(legend: GraphLegend, palette: ChartPalette): string {
  if (legend === "none") {
    return "";
  }

  const position = { top: "t", bottom: "b", left: "l", right: "r" }[legend];

  return `<c:legend><c:legendPos val="${position}"/><c:overlay val="0"/>` +
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>` +
    textPropertiesXml(palette, CHART_TEXT_SIZE) +
    `</c:legend>`;
}

/**
 * Text properties, as both a chart's default and an axis's.
 *
 * Written in the theme's muted ink and body face, so a chart reads as part of
 * the document rather than as something pasted into it — and so that re-theming
 * moves the chart's type with everything else's.
 */
function textPropertiesXml(palette: ChartPalette, size: number): string {
  return `<c:txPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${size}" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0"><a:solidFill><a:srgbClr val="${palette.text}"/></a:solidFill><a:latin typeface="${
    escapeXml(palette.font)
  }"/></a:defRPr></a:pPr><a:endParaRPr lang="en-GB"/></a:p></c:txPr>`;
}

/**
 * How a chart group says its series are arranged.
 *
 * The two words differ by chart: a bar that is not stacked is `clustered`,
 * because its columns stand beside one another, while a line or an area that
 * is not stacked is `standard`, because there is nothing to stand beside.
 * Both stack under the same name, and both share a hundred per cent under the
 * same one.
 *
 * @param node The chart node.
 * @param apart What this chart calls series drawn separately.
 * @returns The `c:grouping` value.
 */
function groupingOf(node: GraphNode, apart: "clustered" | "standard"): string {
  if (node.stacked === "percent") {
    return "percentStacked";
  }

  return node.stacked === true ? "stacked" : apart;
}

/** Whether the series are stacked at all, to a total or to a share. */
function isStacked(node: GraphNode): boolean {
  return node.stacked === true || node.stacked === "percent";
}

/** The plot itself — one chart group, holding every series. */
function plotXml(
  node: GraphNode,
  data: { categories: string[]; series: GraphSeries[] },
  colors: readonly string[],
): string {
  const type = node.graphType;
  const labels = dataLabelsXml(node);
  const axes = `<c:axId val="${CATEGORY_AXIS_ID}"/><c:axId val="${VALUE_AXIS_ID}"/>`;
  const series = data.series
    .map((entry, index) => seriesXml(node, entry, index, data, colors))
    .join("");

  if (type === "pie" || type === "doughnut") {
    // A pie's colours belong to its slices, not to the one series, so
    // `varyColors` is what tells Word the data points differ.
    return `<c:${type}Chart><c:varyColors val="1"/>${series}${labels}<c:firstSliceAng val="0"/>` +
      (type === "doughnut" ? `<c:holeSize val="50"/>` : "") +
      `</c:${type}Chart>`;
  }

  if (type === "line") {
    return `<c:lineChart><c:grouping val="${
      groupingOf(node, "standard")
    }"/><c:varyColors val="0"/>${series}${labels}<c:marker val="1"/>${axes}</c:lineChart>`;
  }

  if (type === "area") {
    return `<c:areaChart><c:grouping val="${
      groupingOf(node, "standard")
    }"/><c:varyColors val="0"/>${series}${labels}${axes}</c:areaChart>`;
  }

  if (type === "scatter") {
    return `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${series}${labels}${axes}</c:scatterChart>`;
  }

  if (type === "radar") {
    // `marker` rather than `filled`: a filled radar hides every series drawn
    // after the first, and the reason to reach for a radar at all is holding
    // several of them against one another. It takes no grouping — a radar
    // cannot stack, because the ring is the scale.
    return `<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>${series}${labels}${axes}</c:radarChart>`;
  }

  if (type === "bubble") {
    // No `c:bubble3D` here, though the schema has a slot for one between the
    // series and `c:bubbleScale`. Measured: Word will not open a file that
    // uses it — not a chart drawn flat, the whole package refused — and Word's
    // own bubble charts leave it out too, stating the flat drawing on each
    // series instead, which is where this states it.
    //
    // `bubbleScale` is a percentage of Word's default area, and 100 is that
    // default. `showNegBubbles` is off because a negative size is not a
    // smaller bubble, it is a figure that does not belong on this chart, and
    // drawing it hollow would present a mistake as a reading.
    return `<c:bubbleChart><c:varyColors val="0"/>${series}${labels}` +
      `<c:bubbleScale val="100"/><c:showNegBubbles val="0"/>` +
      `${axes}</c:bubbleChart>`;
  }

  // Bars, standing or lying. `overlap` is what closes the gap between the
  // segments of a stack; left at the clustered value a stacked bar draws as a
  // staircase.
  const stacked = isStacked(node);

  return `<c:barChart><c:barDir val="${
    type === "barHorizontal" ? "bar" : "col"
  }"/><c:grouping val="${
    groupingOf(node, "clustered")
  }"/><c:varyColors val="0"/>${series}${labels}<c:gapWidth val="${
    stacked ? 60 : 150
  }"/><c:overlap val="${stacked ? 100 : -27}"/>${axes}</c:barChart>`;
}

/** One series, with the values cached so Word draws without opening anything. */
function seriesXml(
  node: GraphNode,
  series: GraphSeries,
  index: number,
  data: { categories: string[]; series: GraphSeries[] },
  colors: readonly string[],
): string {
  const type = node.graphType;
  const rows = data.categories.length;
  const { value: column, size: sizeColumn } = columnsOf(type, index);
  // Past the end of the palette the colours repeat, so a ninth series is drawn
  // the same as the first. That is a bad chart and it is meant to look like
  // one: the alternative is inventing a hue, which puts two colours nobody
  // checked side by side and reads as deliberate. A chart with nine series
  // wants fewer series, or a palette of its own.
  const fill = series.color ?? colors[index % colors.length];
  const head = `<c:idx val="${index}"/><c:order val="${index}"/>` +
    nameXml(series.label, column);

  // A line is drawn, not filled, so its colour belongs on the stroke. Filling
  // a line series paints the area under it, which is a different chart. A
  // radar is a line closed into a ring, and drawn the same way for the same
  // reason.
  const paint = type === "line" || type === "scatter" || type === "radar"
    ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:marker>`
    // A bubble is a disc a reader is meant to see the one behind, so its fill
    // is let down to three quarters. Opaque, a big bubble simply deletes every
    // small one it covers, which is the reading the chart exists to show.
    : type === "bubble"
    ? `<c:spPr><a:solidFill><a:srgbClr val="${fill}"><a:alpha val="75000"/></a:srgbClr></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></a:ln></c:spPr>`
    : `<c:spPr><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>`;

  // A pie has one series and many colours, so each slice is painted as its own
  // data point. Every other chart paints the series and lets the points follow.
  const points = type === "pie" || type === "doughnut"
    ? data.categories
      .map((_, slice) =>
        `<c:dPt><c:idx val="${slice}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${
          colors[slice % colors.length]
        }"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`
      )
      .join("")
    : "";

  const categories = categoriesXml(data.categories, rows);
  const values = valuesXml(series.values, column, rows, node.numberFormat);

  if (type === "scatter") {
    return `<c:ser>${head}${paint}${points}${
      xValuesXml(data.categories, rows)
    }${asYValues(values)}<c:smooth val="0"/></c:ser>`;
  }

  if (type === "bubble") {
    // `invertIfNegative` before the values, and `bubble3D` after them: a
    // bubble series is the one place the schema puts a flag on either side of
    // the numbers, and Word will not open a file that has them the other way.
    return `<c:ser>${head}${paint}<c:invertIfNegative val="0"/>${points}${
      xValuesXml(data.categories, rows)
    }${asYValues(values)}${
      bubbleSizesXml(series.sizes ?? [], sizeColumn, rows)
    }<c:bubble3D val="0"/></c:ser>`;
  }

  if (type === "line" || type === "radar") {
    return `<c:ser>${head}${paint}${points}${categories}${values}` +
      // A radar's ring is already closed, so it has no line to smooth and the
      // schema gives it nowhere to say so.
      (type === "line" ? `<c:smooth val="0"/>` : "") +
      `</c:ser>`;
  }

  if (type === "pie" || type === "doughnut") {
    return `<c:ser>${head}${paint}${points}${categories}${values}</c:ser>`;
  }

  return `<c:ser>${head}${paint}<c:invertIfNegative val="0"/>${points}${categories}${values}</c:ser>`;
}

/**
 * Which spreadsheet columns a series' figures are written in.
 *
 * One column each, except a bubble chart: a bubble series carries two runs of
 * numbers — where the point sits and how big it is drawn — so it takes two
 * columns side by side. Everything that names a cell asks here, so the chart's
 * formulas and the workbook's rows cannot disagree about where a figure lives.
 */
function columnsOf(type: GraphNode["graphType"], index: number): {
  value: string;
  size: string;
} {
  return type === "bubble"
    ? { value: columnName(index * 2 + 1), size: columnName(index * 2 + 2) }
    : { value: columnName(index + 1), size: columnName(index + 1) };
}

/**
 * A series' numbers, renamed as the y of a chart that measures both ways.
 *
 * A scatter and a bubble hold the same cached numbers under `c:yVal` rather
 * than `c:val`, and nothing else about them differs — so the reference is
 * built once and relabelled rather than written twice.
 */
function asYValues(values: string): string {
  return values.replace("<c:val>", "<c:yVal>").replace("</c:val>", "</c:yVal>");
}

/**
 * How big each bubble is drawn, cached against its own column.
 *
 * A size nobody gave is left out rather than written as zero, the same way a
 * missing value is: `dispBlanksAs` is `gap`, and a bubble of no size is a
 * reading that was not taken rather than one that came back as nothing.
 */
function bubbleSizesXml(
  sizes: readonly (number | null)[],
  column: string,
  rows: number,
): string {
  const points = sizes
    .map((size, index) =>
      size === null || size === undefined || !Number.isFinite(size)
        ? ""
        : `<c:pt idx="${index}"><c:v>${size}</c:v></c:pt>`
    )
    .join("");

  return `<c:bubbleSize><c:numRef><c:f>${SHEET}!$${column}$2:$${column}$${
    rows + 1
  }</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${rows}"/>${points}</c:numCache></c:numRef></c:bubbleSize>`;
}

/** A series' name, cached and pointed at the cell the workbook holds it in. */
function nameXml(label: string | undefined, column: string): string {
  if (label === undefined) {
    return "";
  }

  return `<c:tx><c:strRef><c:f>${SHEET}!$${column}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${
    escapeXml(label)
  }</c:v></c:pt></c:strCache></c:strRef></c:tx>`;
}

/** The categories, as the text reference Word reads them from. */
function categoriesXml(categories: readonly string[], rows: number): string {
  const points = categories
    .map((value, index) => `<c:pt idx="${index}"><c:v>${escapeXml(value)}</c:v></c:pt>`)
    .join("");

  return `<c:cat><c:strRef><c:f>${SHEET}!$A$2:$A$${
    rows + 1
  }</c:f><c:strCache><c:ptCount val="${rows}"/>${points}</c:strCache></c:strRef></c:cat>`;
}

/**
 * A scatter's x values, which are numbers rather than names.
 *
 * A category that will not parse counts as its own position, so a scatter
 * given words on its x axis still plots in the order it was written rather
 * than refusing to draw.
 */
function xValuesXml(categories: readonly string[], rows: number): string {
  const points = categories
    .map((value, index) => {
      const parsed = Number.parseFloat(value);

      return `<c:pt idx="${index}"><c:v>${
        Number.isFinite(parsed) ? parsed : index + 1
      }</c:v></c:pt>`;
    })
    .join("");

  return `<c:xVal><c:numRef><c:f>${SHEET}!$A$2:$A$${
    rows + 1
  }</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${rows}"/>${points}</c:numCache></c:numRef></c:xVal>`;
}

/**
 * A series' numbers, cached against the column the workbook writes them in.
 *
 * A `null` is written as no point at all rather than as an empty one: the
 * count still names every row, so Word knows the reading is missing rather
 * than that the series is short.
 */
function valuesXml(
  values: readonly (number | null)[],
  column: string,
  rows: number,
  format: string | undefined,
): string {
  const points = values
    .map((value, index) =>
      value === null || value === undefined || !Number.isFinite(value)
        ? ""
        : `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`
    )
    .join("");

  return `<c:val><c:numRef><c:f>${SHEET}!$${column}$2:$${column}$${
    rows + 1
  }</c:f><c:numCache><c:formatCode>${
    escapeXml(format ?? "General")
  }</c:formatCode><c:ptCount val="${rows}"/>${points}</c:numCache></c:numRef></c:val>`;
}

/**
 * Whether each point prints its own figure.
 *
 * Where the label sits is left unsaid on purpose: the position a chart allows
 * depends on its type — a stacked bar takes `ctr` and a clustered one `outEnd`
 * — and a position the type does not allow is a file Word offers to repair.
 * Saying nothing takes the default, which is the right one for every type.
 */
function dataLabelsXml(node: GraphNode): string {
  if (node.dataLabels !== true) {
    return "";
  }

  const percent = node.graphType === "pie" || node.graphType === "doughnut";

  return `<c:dLbls>${
    node.numberFormat === undefined
      ? ""
      : `<c:numFmt formatCode="${escapeXml(node.numberFormat)}" sourceLinked="0"/>`
  }<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:showLegendKey val="0"/><c:showVal val="${
    percent ? 0 : 1
  }"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${
    percent ? 1 : 0
  }"/><c:showBubbleSize val="0"/></c:dLbls>`;
}

/** The axes a chart type has, ruled in the theme's own hairline. */
function axesXml(node: GraphNode, palette: ChartPalette): string {
  if (node.graphType === "pie" || node.graphType === "doughnut") {
    return "";
  }

  // Gridlines run across the value axis, so on a horizontal bar chart — where
  // the value axis is the one along the bottom — they run up and down. Word
  // works that out from `axPos`, which is the only thing that changes.
  const horizontal = node.graphType === "barHorizontal";
  const value = valueAxisXml(node, palette, horizontal ? "b" : "l", horizontal);
  // A bubble measures along both axes for the same reason a scatter does: its
  // x is a figure, not a name, and a category axis would space the points
  // evenly and throw that figure away.
  const measured = node.graphType === "scatter" || node.graphType === "bubble";
  const category = measured
    ? scatterCategoryAxisXml(node, palette, horizontal ? "l" : "b")
    : categoryAxisXml(
      node,
      palette,
      horizontal ? "l" : "b",
      horizontal,
      node.graphType === "radar",
    );

  // The category axis is written first because that is the order Word writes
  // it in, and a chart is easier to diff against one Word produced.
  return category + value;
}

/**
 * The axis the categories are counted along.
 *
 * A horizontal bar chart counts its categories the other way up. OOXML's
 * default puts the first one at the bottom — which is what Word does when you
 * insert a bar chart from a table, and what a document written here should not
 * do: the categories were given as an array, and the first entry of an array
 * belongs at the top. So the scale is reversed, and the value axis is told to
 * cross at the far end so the labels stay on the left where they were.
 *
 * `webbed` is the radar's spokes. Every other chart rules its grid across the
 * value axis alone — a line up from each category would be a second grid
 * saying what the tick marks already say — but a radar has no plot area to
 * rule, and the spokes out to each category are what makes the ring readable
 * as a ring rather than as a shape floating in space.
 */
function categoryAxisXml(
  node: GraphNode,
  palette: ChartPalette,
  position: string,
  reversed = false,
  webbed = false,
): string {
  return `<c:catAx><c:axId val="${CATEGORY_AXIS_ID}"/><c:scaling><c:orientation val="${
    reversed ? "maxMin" : "minMax"
  }"/></c:scaling><c:delete val="0"/><c:axPos val="${position}"/>` +
    (webbed ? gridlinesXml(palette) : "") +
    axisTitleXml(node.categoryAxisTitle, palette, position === "l") +
    `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>` +
    axisLineXml(palette) +
    textPropertiesXml(palette, CHART_TEXT_SIZE) +
    `<c:crossAx val="${VALUE_AXIS_ID}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
}

/** A scatter's x axis, which measures rather than names. */
function scatterCategoryAxisXml(
  node: GraphNode,
  palette: ChartPalette,
  position: string,
): string {
  return `<c:valAx><c:axId val="${CATEGORY_AXIS_ID}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${position}"/>` +
    axisTitleXml(node.categoryAxisTitle, palette, position === "l") +
    `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>` +
    axisLineXml(palette) +
    textPropertiesXml(palette, CHART_TEXT_SIZE) +
    `<c:crossAx val="${VALUE_AXIS_ID}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`;
}

/**
 * The axis the values are measured against, and the one carrying the grid.
 *
 * `crossing` is where it meets the category axis. It stays at zero everywhere
 * except a horizontal bar chart, whose category axis has been turned upside
 * down so the first category reads first — and where a value axis still
 * crossing at zero would take the category labels over to the right-hand side
 * with it.
 */
function valueAxisXml(
  node: GraphNode,
  palette: ChartPalette,
  position: string,
  crossesAtEnd = false,
): string {
  const format = valueAxisFormat(node);

  return `<c:valAx><c:axId val="${VALUE_AXIS_ID}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${position}"/>` +
    gridlinesXml(palette) +
    axisTitleXml(node.valueAxisTitle, palette, position === "l") +
    `<c:numFmt formatCode="${escapeXml(format ?? "General")}" sourceLinked="${
      format === undefined ? 1 : 0
    }"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>` +
    // The value axis draws no line of its own: the gridlines already carry the
    // scale across the plot, and a rule under them is a second one saying the
    // same thing.
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>` +
    textPropertiesXml(palette, CHART_TEXT_SIZE) +
    `<c:crossAx val="${CATEGORY_AXIS_ID}"/><c:crosses val="${
      crossesAtEnd ? "max" : "autoZero"
    }"/><c:crossBetween val="between"/></c:valAx>`;
}

/**
 * What the value axis prints its numbers in.
 *
 * The document's format where it named one. Where it did not, a percent stack
 * takes `0%`, because a percent stack's axis is a share and its ticks run to
 * one rather than to a total — printed unformatted, a 100% stacked chart is
 * labelled 0, 0.2, 0.4, which is a scale nobody means. Every other chart says
 * nothing and takes the format the numbers already carry.
 */
function valueAxisFormat(node: GraphNode): string | undefined {
  if (node.numberFormat !== undefined) {
    return node.numberFormat;
  }

  return node.stacked === "percent" ? "0%" : undefined;
}

/** The grid a chart is ruled with, in the theme's own hairline. */
function gridlinesXml(palette: ChartPalette): string {
  return `<c:majorGridlines><c:spPr><a:ln w="${GRIDLINE_WIDTH_EMU}" cap="flat"><a:solidFill><a:srgbClr val="${palette.rule}"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>`;
}

/** The hairline an axis is ruled in, in the theme's rule colour. */
function axisLineXml(palette: ChartPalette): string {
  return `<c:spPr><a:noFill/><a:ln w="${GRIDLINE_WIDTH_EMU}" cap="flat"><a:solidFill><a:srgbClr val="${palette.rule}"/></a:solidFill><a:round/></a:ln></c:spPr>`;
}

/** What an axis is labelled, turned on its side where the axis is vertical. */
function axisTitleXml(
  title: string | undefined,
  palette: ChartPalette,
  vertical: boolean,
): string {
  if (title === undefined) {
    return "";
  }

  // A title beside a vertical axis reads bottom-to-top, which OOXML says as a
  // rotation of -60° in sixtieths of a degree.
  const rotation = vertical ? ` rot="-5400000" vert="horz"` : ` rot="0" vert="horz"`;

  return `<c:title><c:tx><c:rich><a:bodyPr${rotation}/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${CHART_TEXT_SIZE}" b="0"><a:solidFill><a:srgbClr val="${palette.text}"/></a:solidFill><a:latin typeface="${
    escapeXml(palette.font)
  }"/></a:defRPr></a:pPr><a:r><a:t>${
    escapeXml(title)
  }</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:title>`;
}

// ---------------------------------------------------------------------------
// The workbook
//
// The smallest spreadsheet that is still a spreadsheet: a workbook, one sheet,
// and the rows. No styles, no shared strings — every label is written into the
// cell that holds it, because a string table for a dozen words is a second part
// to keep in step with the first for nothing.
// ---------------------------------------------------------------------------

/**
 * The chart's numbers as a workbook, one part per file it needs.
 *
 * The columns are the ones {@linkcode columnsOf} named, so what "Edit Data"
 * opens is laid out exactly where the chart's own formulas say it is. A bubble
 * takes two columns per series; every other chart takes one.
 */
function workbookOf(
  type: GraphNode["graphType"],
  data: { categories: string[]; series: GraphSeries[] },
): ZipEntry[] {
  const bubble = type === "bubble";
  const header = data.series
    .map((series, index) => {
      const columns = columnsOf(type, index);

      if (series.label === undefined) {
        return "";
      }

      // A bubble's second column is headed too. Unheaded, a reader opening the
      // data finds a column of figures with nothing saying what they weigh.
      return cell(`${columns.value}1`, series.label) +
        (bubble ? cell(`${columns.size}1`, `${series.label} (size)`) : "");
    })
    .join("");

  const rows = [
    `<row r="1">${header}</row>`,
    ...data.categories.map((category, row) =>
      `<row r="${row + 2}">${cell(`A${row + 2}`, category)}${
        data.series
          .map((series, index) => {
            const columns = columnsOf(type, index);

            return numberCell(columns.value, row, series.values[row]) +
              (bubble ? numberCell(columns.size, row, series.sizes?.[row]) : "");
          })
          .join("")
      }</row>`
    ),
  ].join("");

  return [
    part(
      "[Content_Types].xml",
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    ),
    part(
      "_rels/.rels",
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    ),
    part(
      "xl/workbook.xml",
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${SHEET}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    part(
      "xl/_rels/workbook.xml.rels",
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
    part(
      "xl/worksheets/sheet1.xml",
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`,
    ),
  ];
}

/**
 * One numeric cell, or nothing at all where the reading is missing.
 *
 * An empty cell rather than a zero, which is the same answer the chart's cache
 * gives: a row nobody filled in should look unfilled when the data is opened.
 */
function numberCell(column: string, row: number, value: number | null | undefined): string {
  return value === null || value === undefined || !Number.isFinite(value)
    ? ""
    : `<c r="${column}${row + 2}"><v>${value}</v></c>`;
}

/** One text cell, written in place rather than through a string table. */
function cell(reference: string, value: string): string {
  return `<c r="${reference}" t="inlineStr"><is><t xml:space="preserve">${
    escapeXml(value)
  }</t></is></c>`;
}

/** One part of the workbook, with the declaration every OOXML part carries. */
function part(name: string, xml: string): ZipEntry {
  return {
    name,
    bytes: new TextEncoder().encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + xml,
    ),
  };
}

/** A spreadsheet column's name, counted from zero: A, B, … Z, AA, AB. */
function columnName(index: number): string {
  let name = "";
  let remaining = index;

  do {
    name = String.fromCharCode(65 + (remaining % 26)) + name;
    remaining = Math.floor(remaining / 26) - 1;
  } while (remaining >= 0);

  return name;
}

/** Text as XML says it, so a label with an ampersand in it is still a label. */
function escapeXml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}
