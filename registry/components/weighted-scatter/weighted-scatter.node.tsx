import {
  Graph,
  Paragraph,
  Section,
  type Section as SectionComponent,
  useFormat,
  useState,
} from "docxcelerate/template";

/**
 * Two measures against each other, with a third as the weight — and the
 * sentence naming what sits in the corner.
 *
 * The chart a scatter turns into when "which of these matters" is the real
 * question. Plotting cost against usage tells you where things sit; sizing the
 * points by how many people it affects tells you which of them is worth the
 * page. A bubble is the only chart here that carries a third figure without
 * inventing a second axis to hang it on.
 *
 * **The size is an area, and the component keeps it one.** A bubble drawn with
 * its *radius* proportional to the figure shows twice the reading as four
 * times the ink, which is the single most common way this chart lies. Word
 * scales by area, and `weight` is passed through untouched so it stays that
 * way — do not pre-square it.
 *
 * Three more decisions come with it.
 *
 * **A point with no weight is not drawn.** A bubble of no size is a reading
 * nobody took, not a reading of nothing, and drawing it as a dot puts a point
 * on the chart that has no third figure while every other point does.
 *
 * **Labels are in the prose, not on the points.** A bubble chart labelled
 * point by point is unreadable at a text column's width the moment two bubbles
 * touch. The chart shows the shape of the field; the sentence names the corner
 * worth looking at, and the key names the series.
 *
 * **The corner is named, not left to the reader.** Which point is furthest out
 * is worked out here from the same rows the chart is drawn from, so the prose
 * and the plot cannot disagree about which one it is.
 *
 * Installed by `dxcl add weighted-scatter`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface WeightedScatterData {
  scatter: {
    /** What the horizontal axis measures: `"Cost per visit"`. */
    xLabel: string;
    /** What the vertical axis measures: `"Visits per week"`. */
    yLabel: string;
    /** What the size of a point means: `"members affected"`. For the caption. */
    weightLabel: string;
    /**
     * One entry per point, in any order.
     *
     * A scatter has no sequence of its own, so nothing here is sorted for
     * drawing — only for finding the point the sentence names.
     */
    points: Array<{
      /** What this point is, for the sentence. Never drawn on the plot. */
      label: string;
      /** Where it sits along the horizontal axis. */
      x: number;
      /** Where it sits up the vertical one. */
      y: number;
      /**
       * How much it weighs, drawn as the point's area.
       *
       * Passed to Word untouched, which scales by area. Do not square it
       * first — a bubble whose radius is the figure shows twice the reading
       * as four times the ink.
       */
      weight: number;
    }>;
    /** How the axes print their figures, as an OOXML number format. */
    numberFormat?: string;
  };
}

export const WeightedScatter: SectionComponent = () => {
  const format = useFormat();
  const [scatter] = useState((data: WeightedScatterData) => data.scatter);

  // A point missing any of its three figures is not a point on this chart. A
  // weight of nothing especially: drawn, it is a dot with no third reading
  // sitting among bubbles that all have one.
  const points = (scatter.points ?? []).filter((point) =>
    Number.isFinite(point.x) && Number.isFinite(point.y) &&
    Number.isFinite(point.weight) && point.weight > 0
  );

  if (points.length === 0) {
    return (
      <Section id="weighted-scatter" title={`${scatter.yLabel} against ${scatter.xLabel}`}>
        <Paragraph id="weighted-scatter-none">
          We have nothing recorded to plot here yet.
        </Paragraph>
      </Section>
    );
  }

  const heaviest = [...points].sort((left, right) => right.weight - left.weight)[0];
  const dropped = (scatter.points ?? []).length - points.length;

  return (
    <Section id="weighted-scatter" title={`${scatter.yLabel} against ${scatter.xLabel}`}>
      <Paragraph id="weighted-scatter-summary">
        {format.number(points.length)} {format.plural(points.length, "point")} plotted, sized by{" "}
        {scatter.weightLabel}. The largest is {heaviest.label}, at{" "}
        {format.number(heaviest.weight)} {scatter.weightLabel} —{" "}
        {scatter.xLabel.toLowerCase()} {format.number(heaviest.x)},{" "}
        {scatter.yLabel.toLowerCase()} {format.number(heaviest.y)}.
        {dropped > 0
          ? ` ${format.number(dropped)} ${format.plural(dropped, "entry", "entries")} had no ` +
            `${scatter.weightLabel} recorded and ${
              format.plural(dropped, "is", "are")
            } not shown.`
          : ""}
      </Paragraph>

      <Graph
        id="weighted-scatter-chart"
        graphType="bubble"
        // The categories are the x values — a bubble measures along both axes,
        // so these are numbers written as text rather than names.
        categoryAxisTitle={scatter.xLabel}
        valueAxisTitle={scatter.yLabel}
        numberFormat={scatter.numberFormat}
        // One series, so a key would name the chart's own title back at the
        // reader. What each bubble is belongs in the prose, not on the plot.
        legend="none"
        data={{
          categories: points.map((point) => String(point.x)),
          series: [
            {
              label: scatter.yLabel,
              values: points.map((point) => point.y),
              // Straight through. Word draws the area from this figure.
              sizes: points.map((point) => point.weight),
            },
          ],
        }}
        caption={`Each bubble is one entry; its size is ${scatter.weightLabel}.`}
      />
    </Section>
  );
};
