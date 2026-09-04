import {
  Graph,
  Paragraph,
  Section,
  type Section as SectionComponent,
  useFormat,
  useState,
} from "docxcelerate/template";

/**
 * This period against the last, category by category.
 *
 * The chart every report reaches for, and the one that goes wrong in the same
 * two ways every time.
 *
 * **It turns its bars on their side when the labels are long.** Category names
 * coming out of a real system are "Ground floor maintenance", not "Q1", and
 * under a vertical bar those either overlap, shrink, or tilt forty-five
 * degrees and become something the reader has to lean over to read. Laid down,
 * the label sits beside its bar at full size and the chart grows downwards,
 * which is the direction a page has room in. The threshold is on the longest
 * label rather than on the number of rows, because that is what actually
 * decides whether the axis fits.
 *
 * **It never puts two measures on one chart.** Both series here are the same
 * measure at two different times, which is the one comparison a shared axis is
 * honest about. If you find yourself wanting spend on one axis and headcount
 * on the other, that is two charts — a second axis makes the crossing point
 * look like a finding, and it is an artefact of the scales you chose.
 *
 * The sentence above the chart names what moved and in which direction, worked
 * out from the same rows, so the prose cannot drift from the plot.
 *
 * Installed by `dxcl add period-comparison`.
 */

/**
 * How long a label may be before the bars are laid on their side.
 *
 * Sixteen characters, which is about where a label stops fitting under a bar
 * on a text column's worth of chart at eleven point.
 */
const LONG_LABEL = 16;

/** What this component reads. Add these fields to your document data type. */
export interface PeriodComparisonData {
  comparison: {
    /** What the earlier run is called: `"2024"`, `"Last year"`. */
    previousLabel: string;
    /** What the later run is called: `"2025"`, `"This year"`. */
    currentLabel: string;
    /**
     * One row per category, in the order they should be drawn.
     *
     * A `null` is a category that was not measured in that period, and draws
     * as a gap rather than as a bar of nothing.
     */
    rows: Array<{
      /** What the category is called, as it should be printed. */
      label: string;
      /** The earlier figure. */
      previous: number | null;
      /** The later figure. */
      current: number | null;
    }>;
    /** What the figures measure — printed beside the value axis. */
    measure?: string;
    /** What the categories are — printed beside the category axis. */
    dimension?: string;
    /** How the value axis prints its figures, as an OOXML number format. */
    numberFormat?: string;
  };
}

export const PeriodComparison: SectionComponent = () => {
  const format = useFormat();
  const [comparison] = useState((data: PeriodComparisonData) => data.comparison);

  const rows = comparison.rows ?? [];
  // Long labels do not fit under a vertical bar. Measured on the longest,
  // because one runaway name is enough to wreck the axis for all of them.
  const longest = rows.reduce((most, row) => Math.max(most, row.label.length), 0);

  // Only the rows measured in both periods can be said to have moved. One
  // measured in neither is not news; one measured in only one is a gap the
  // chart shows and the sentence has no honest comparison for.
  const moved = rows.filter(
    (row): row is { label: string; previous: number; current: number } =>
      typeof row.previous === "number" && typeof row.current === "number",
  );
  const risen = moved.filter((row) => row.current > row.previous).map((row) => row.label);
  const fallen = moved.filter((row) => row.current < row.previous).map((row) => row.label);

  if (rows.length === 0) {
    return (
      <Section id="period-comparison" title="This period against the last">
        <Paragraph id="period-comparison-none">
          We have nothing to compare for this period yet.
        </Paragraph>
      </Section>
    );
  }

  return (
    <Section id="period-comparison" title="This period against the last">
      <Paragraph id="period-comparison-summary">
        {risen.length === 0 && fallen.length === 0
          ? `Nothing changed between ${comparison.previousLabel} and ` +
            `${comparison.currentLabel}.`
          : [
            `Against ${comparison.previousLabel}, `,
            risen.length > 0 ? `${format.list(risen)} rose` : "",
            risen.length > 0 && fallen.length > 0 ? " and " : "",
            fallen.length > 0 ? `${format.list(fallen)} fell` : "",
            ".",
          ].join("")}
      </Paragraph>

      <Graph
        id="period-comparison-chart"
        title={`${comparison.currentLabel} against ${comparison.previousLabel}`}
        graphType={longest > LONG_LABEL ? "barHorizontal" : "bar"}
        // Two series, so the key names which is which. Without it the chart
        // has two colours and nothing saying what either one is.
        legend="bottom"
        numberFormat={comparison.numberFormat ?? "#,##0"}
        categoryAxisTitle={comparison.dimension}
        valueAxisTitle={comparison.measure}
        data={{
          categories: rows.map((row) => row.label),
          series: [
            {
              label: comparison.previousLabel,
              values: rows.map((row) => row.previous ?? null),
            },
            {
              label: comparison.currentLabel,
              values: rows.map((row) => row.current ?? null),
            },
          ],
        }}
        caption={`${comparison.currentLabel} beside ${comparison.previousLabel}, ` +
          `by ${comparison.dimension ?? "category"}.`}
      />
    </Section>
  );
};
