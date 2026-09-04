import * as echarts from "echarts";
import type { ChartDrawer, PreviewChart, PreviewChartSeries } from "docxcelerate/preview";

/**
 * Drawing the preview's charts with ECharts.
 *
 * The framework packs a chart as a real Word chart — a `c:chartSpace` part with
 * the numbers in it — and Word draws it. Nothing in the browser can draw what
 * Word draws, and docx-preview does not try: it leaves an empty frame at
 * exactly the size the file gave the chart, which is what keeps the preview
 * paginating like Word.
 *
 * This fills that frame. Everything it plots is read back out of the packed
 * file rather than out of the document that produced it — the type, the
 * numbers, the colours the theme chose, the format the value axis prints — so
 * what is on screen is a drawing of what is in the `.docx` and not a second
 * opinion about what should have been.
 *
 * It is a drawing all the same, and not a screenshot of Word. Bars will be a
 * pixel or two from where Word puts them and a label may wrap differently.
 * What is exact is the frame, which is the part the page is laid out around.
 * Open the `.docx` when the plot itself has to be right.
 *
 * ECharts is used in server-side mode even here, because that is the mode that
 * returns markup rather than taking over an element: the frame already exists
 * and already has its size, so there is nothing to attach to and nothing to
 * measure.
 */

/** A point, in the CSS pixels a browser draws in. */
const PX_PER_PT = 96 / 72;

/** How much room the axis labels are given, in pixels. */
const GUTTER = { left: 56, right: 18, top: 16, bottom: 18 };

/**
 * A drawer for {@linkcode settleDocxPreviewCharts}.
 *
 * @returns A function that turns one chart read out of the file into SVG.
 */
export function createChartDrawer(): ChartDrawer {
  return (chart) => {
    const instance = echarts.init(null, null, {
      renderer: "svg",
      ssr: true,
      // The frame's own size, in the pixels its `pt` width comes to. Anything
      // else and the plot is drawn for a box it is not in.
      width: Math.round(chart.widthPt * PX_PER_PT),
      height: Math.round(chart.heightPt * PX_PER_PT),
    });

    instance.setOption(optionFor(chart));
    const svg = instance.renderToSVGString();
    instance.dispose();

    return svg;
  };
}

/** What ECharts is asked to draw, from what the file said. */
function optionFor(chart: PreviewChart): echarts.EChartsOption {
  const ink = hex(chart.textColor) ?? "#595959";
  const rule = hex(chart.ruleColor) ?? "#D9D9D9";
  const font = chart.font ?? "inherit";
  const pie = chart.graphType === "pie" || chart.graphType === "doughnut";
  const horizontal = chart.graphType === "barHorizontal";
  // Whether there is anything for a key to name. A pie is keyed by its
  // slices, so it has one whenever it has categories; everything else is keyed
  // by its series, and a series with no name would put an empty row in it.
  const named = pie
    ? chart.categories.length > 0
    : chart.series.some((series) => series.label !== undefined);

  const text = { color: ink, fontFamily: font, fontSize: 11 };
  const values = valueAxis(rule, text, chart.numberFormat);
  // A scatter counts along numbers on both axes, so what stands where the
  // categories would is a second measured axis rather than a named one.
  const categories = chart.graphType === "scatter"
    ? valueAxis(rule, text)
    : categoryAxis(chart, rule, text);

  return {
    // A preview is a still picture of a document. A chart that animated into
    // place would be a page that moves under a reader who is checking a
    // layout, and it would be caught mid-flight by anything screenshotting it.
    animation: false,
    textStyle: { fontFamily: font, color: ink },
    ...(chart.title === undefined ? {} : {
      title: {
        text: chart.title,
        left: "center",
        top: 4,
        textStyle: { ...text, fontSize: 13, fontWeight: "normal" as const },
      },
    }),
    ...(chart.legend === "none" || !named ? {} : { legend: legendFor(chart, text) }),
    ...(pie ? {} : {
      grid: {
        left: (horizontal ? GUTTER.left + 20 : GUTTER.left) +
          (chart.valueAxisTitle === undefined && chart.categoryAxisTitle === undefined ? 0 : 16),
        right: GUTTER.right,
        top: GUTTER.top + (chart.title === undefined ? 0 : 22),
        bottom: GUTTER.bottom + (chart.legend === "none" || !named ? 0 : 22) +
          (chart.categoryAxisTitle === undefined && chart.valueAxisTitle === undefined ? 0 : 18),
        containLabel: true,
      },
      // The value axis is the one along the bottom on a horizontal bar chart,
      // so the labels follow the axes rather than the names x and y.
      xAxis: horizontal
        ? { ...values, ...axisTitle(chart.valueAxisTitle, text, false) }
        : { ...categories, ...axisTitle(chart.categoryAxisTitle, text, false) },
      // `inverse` puts the first category at the top, which is where the file
      // puts it: the packer writes the category axis `maxMin` for a horizontal
      // bar so it reads in the order the array was written. ECharts, like
      // Word's own bar charts, would otherwise start from the bottom — and a
      // preview that reads bottom-up while the `.docx` reads top-down is a
      // preview disagreeing with the document it was made from.
      yAxis: horizontal
        ? { ...categories, inverse: true, ...axisTitle(chart.categoryAxisTitle, text, true) }
        : { ...values, ...axisTitle(chart.valueAxisTitle, text, true) },
    }),
    series: pie ? [pieSeries(chart)] : chart.series.map((series) => plotSeries(chart, series)),
  };
}

/** How every axis is drawn: recessive, in the theme's own rule colour and ink. */
type AxisText = { color: string; fontFamily: string; fontSize: number };

/**
 * The axis the values are measured against, carrying the grid.
 *
 * The grid runs across this axis and not the other, which is what the packed
 * chart says too: `c:majorGridlines` is written on the value axis alone.
 */
function valueAxis(rule: string, text: AxisText, format?: string) {
  return {
    type: "value" as const,
    axisLine: { lineStyle: { color: rule } },
    axisTick: { show: false },
    axisLabel: {
      ...text,
      ...(format === undefined ? {} : { formatter: (value: number) => print(value, format) }),
    },
    splitLine: { lineStyle: { color: rule } },
  };
}

/** The axis the categories are named along, which carries no grid of its own. */
function categoryAxis(chart: PreviewChart, rule: string, text: AxisText) {
  return {
    type: "category" as const,
    data: [...chart.categories],
    axisLine: { lineStyle: { color: rule } },
    axisTick: { show: false },
    axisLabel: text,
    splitLine: { show: false },
  };
}

/** What an axis is labelled, where the file gives it a label. */
function axisTitle(title: string | undefined, text: AxisText, vertical: boolean) {
  if (title === undefined) {
    return {};
  }

  return {
    name: title,
    // Beside the axis rather than at its end, and turned on its side where the
    // axis is vertical — which is what Word does with the same label.
    nameLocation: "middle" as const,
    nameGap: vertical ? 44 : 26,
    nameRotate: vertical ? 90 : 0,
    nameTextStyle: text,
  };
}

/** Where the key sits, said the way ECharts says it. */
function legendFor(
  chart: PreviewChart,
  text: { color: string; fontFamily: string; fontSize: number },
): echarts.LegendComponentOption {
  const place = {
    top: { top: chart.title === undefined ? 4 : 26, left: "center" as const },
    bottom: { bottom: 4, left: "center" as const },
    left: { left: 4, top: "middle" as const, orient: "vertical" as const },
    right: { right: 4, top: "middle" as const, orient: "vertical" as const },
    none: {},
  }[chart.legend];

  // A pie's key names its slices; every other chart's names its series. The
  // packed file says the same thing by putting a pie's colours on its data
  // points rather than on its one series.
  const pie = chart.graphType === "pie" || chart.graphType === "doughnut";

  return {
    ...place,
    itemWidth: 10,
    itemHeight: 10,
    textStyle: text,
    data: pie ? [...chart.categories] : chart.series
      .map((series) => series.label)
      .filter((label): label is string => label !== undefined),
  };
}

/** One plotted series, whichever kind of plot it belongs to. */
function plotSeries(chart: PreviewChart, series: PreviewChartSeries): echarts.SeriesOption {
  const color = hex(series.color);
  const data = series.values.map((value) => value ?? null);
  const stack = chart.stacked ? "total" : undefined;

  if (chart.graphType === "scatter") {
    return {
      name: series.label,
      type: "scatter",
      symbolSize: 8,
      itemStyle: { color },
      data: data.map((value, index) => [Number(chart.categories[index] ?? index + 1), value]),
    };
  }

  if (chart.graphType === "line" || chart.graphType === "area") {
    return {
      name: series.label,
      type: "line",
      stack,
      symbol: "circle",
      symbolSize: 6,
      lineStyle: { color, width: 2 },
      itemStyle: { color },
      // An area chart is a line with the ground filled in under it, which is
      // exactly what OOXML's `c:areaChart` is beside its `c:lineChart`.
      ...(chart.graphType === "area" ? { areaStyle: { color, opacity: 0.25 } } : {}),
      data,
    };
  }

  return {
    name: series.label,
    type: "bar",
    stack,
    itemStyle: { color },
    ...labelStyle(chart),
    data,
  };
}

/**
 * Whether the points on a plot print their own figures.
 *
 * Off for most charts, because a number on every point is a table drawn badly
 * — but it is the document's decision, and the packed file records it.
 */
function labelStyle(chart: PreviewChart) {
  if (chart.dataLabels === "none") {
    return {};
  }

  const format = chart.numberFormat;

  return {
    label: {
      show: true,
      position: "top" as const,
      color: hex(chart.textColor) ?? "#595959",
      fontSize: 10,
      // ECharts hands a label callback the whole data point, whose `value` is
      // anything a series may hold. Only a number can be formatted, and
      // anything else prints as it stands.
      ...(format === undefined ? {} : {
        formatter: (point: { value?: unknown }) =>
          typeof point.value === "number" ? print(point.value, format) : String(point.value ?? ""),
      }),
    },
  };
}

/** A pie or a doughnut: one series, and a colour for every slice. */
function pieSeries(chart: PreviewChart): echarts.SeriesOption {
  const values = chart.series[0]?.values ?? [];

  return {
    type: "pie",
    // A doughnut is a pie with its middle taken out, which is the one thing
    // `c:holeSize` says.
    radius: chart.graphType === "doughnut" ? ["45%", "70%"] : "70%",
    // A pie is the one chart where a figure on every slice is the reading
    // rather than clutter, and the file says whether the document asked for it.
    label: chart.dataLabels === "none"
      ? { show: false }
      : {
        show: true,
        formatter: chart.dataLabels === "percent" ? "{d}%" : "{c}",
        color: hex(chart.textColor) ?? "#595959",
        fontSize: 11,
      },
    data: chart.categories.map((name, index) => ({
      name,
      value: values[index] ?? 0,
      itemStyle: { color: hex(chart.sliceColors[index]) },
    })),
  };
}

/**
 * A number printed the way the value axis says to print it.
 *
 * Enough of an OOXML number format to be useful and no more: whatever stands
 * before the first `#` or `0` is a prefix, whatever stands after the last is a
 * suffix, a comma means thousands are grouped, and the digits after the point
 * say how many decimals. A format this does not understand prints the number
 * plainly, which is what the axis would have shown anyway.
 */
function print(value: number, format: string): string {
  const percent = format.includes("%");
  const digits = /\.(0+)/.exec(format)?.[1].length ?? 0;
  const shown = percent ? value * 100 : value;
  const body = shown.toLocaleString(undefined, {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
    useGrouping: format.includes(","),
  });

  const prefix = /^[^#0]*/.exec(format)?.[0].replace(/[\\"]/g, "") ?? "";
  const suffix = percent ? "%" : "";

  return `${prefix}${body}${suffix}`;
}

/** A colour as the file writes it — six hex digits — as CSS wants it. */
function hex(value: string | undefined): string | undefined {
  return value === undefined ? undefined : `#${value}`;
}
