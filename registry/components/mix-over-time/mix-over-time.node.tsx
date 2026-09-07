import {
  Graph,
  Paragraph,
  Section,
  type Section as SectionComponent,
  useFormat,
  useState,
} from "docxcelerate/template";

/**
 * How a mix shifted, as a hundred per cent stack and as the sentence naming
 * what shifted.
 *
 * A stacked chart answers one of two questions and cannot answer both. Stacked
 * to a total, the height of each column is the size of the period and the
 * segments are its parts. Stacked to a hundred per cent, every column is the
 * same height and what is read is the *mix* — which is the question anybody
 * asking "has this changed?" is actually asking. This component takes the
 * second, and says so on the axis and in the caption, because a reader who
 * thinks they are looking at volumes has been misled by a chart that is
 * otherwise correct.
 *
 * Four decisions are already made here.
 *
 * **The totals are given away, and the prose gives them back.** That is the
 * trade a percent stack makes: a column that halved and a column that doubled
 * are drawn identically. So the sentence carries the first and last totals,
 * which is the fact the chart deliberately dropped.
 *
 * **The raw figures go into the file.** Word works each column's shares out
 * itself and labels the axis 0% to 100%. A component that divided first would
 * hand Word shares of shares and draw a chart of nothing.
 *
 * **The tail is folded into Other, per period rather than per row.** A mix out
 * of real data has as many categories as the system had rows, and past the
 * palette the ninth band is painted the same as the first. Which categories
 * are large is decided once, across the whole span — folding per period would
 * put a category in Other in one column and beside it in the next, which draws
 * as a band that appears and disappears for no reason a reader can see.
 *
 * **A period with nothing recorded is dropped, not drawn empty.** A column
 * with no readings has no mix, and stacked to a hundred per cent an empty
 * column is either a gap in the run or, worse, drawn full of whatever rounding
 * produced.
 *
 * Installed by `dxcl add mix-over-time`.
 */

/**
 * How many bands are drawn before the rest become "Other".
 *
 * Five, one fewer than a pie gets: a band is read across the whole width of
 * the chart rather than as a wedge, so two similar hues are held apart at more
 * points and are harder to tell apart, not easier.
 */
const BANDS = 5;

/** What this component reads. Add these fields to your document data type. */
export interface MixOverTimeData {
  mix: {
    /**
     * One entry per period, oldest first.
     *
     * The order is the chart's order. Sort before it reaches the document — a
     * chart is drawn in the order it is given, and a component that sorted for
     * you would disagree with the table beside it.
     */
    periods: Array<{
      /** What the period is called: `"Jan"`, `"Q1"`, `"2024"`. */
      label: string;
      /**
       * What each category came to in that period.
       *
       * Keyed by category name, and the keys need not agree between periods —
       * a category absent from one period counts as nothing in it, which is
       * what a mix means by absent. Negatives are dropped: a share of a whole
       * cannot be less than none of it.
       */
      amounts: Record<string, number>;
    }>;
    /** What is being counted: `"visits"`, `"tickets"`, `"kWh"`. Printed in the prose. */
    unit?: string;
  };
}

/** One category, as it is drawn: a name and a figure per period. */
interface Band {
  label: string;
  values: number[];
}

export const MixOverTime: SectionComponent = () => {
  const format = useFormat();
  const [mix] = useState((data: MixOverTimeData) => data.mix);

  // A period nobody recorded anything against has no mix to draw. Dropped
  // rather than drawn, because stacked to a hundred per cent an empty column
  // is not an empty column — it is whatever the rounding produced.
  const periods = (mix.periods ?? []).filter((period) =>
    Object.values(period.amounts ?? {}).some((value) => Number.isFinite(value) && value > 0)
  );

  const unit = mix.unit === undefined ? "" : ` ${mix.unit}`;

  if (periods.length < 2) {
    return (
      <Section id="mix-over-time" title="How the mix has changed">
        <Paragraph id="mix-over-time-none">
          We do not have enough history to show how this has changed yet.
        </Paragraph>
      </Section>
    );
  }

  // Which categories are large is decided once, over the whole span. Deciding
  // per period would move a category in and out of Other from one column to
  // the next, which draws as a band that appears and disappears.
  const totals = new Map<string, number>();

  for (const period of periods) {
    for (const [label, amount] of Object.entries(period.amounts ?? {})) {
      if (Number.isFinite(amount) && amount > 0) {
        totals.set(label, (totals.get(label) ?? 0) + amount);
      }
    }
  }

  const ranked = [...totals.entries()].sort((left, right) => right[1] - left[1]);
  const named = ranked.slice(0, BANDS).map(([label]) => label);
  const folded = ranked.length > named.length;

  const amountOf = (period: (typeof periods)[number], label: string) => {
    const value = period.amounts?.[label];

    return Number.isFinite(value) && value > 0 ? value : 0;
  };

  const bands: Band[] = named.map((label) => ({
    label,
    values: periods.map((period) => amountOf(period, label)),
  }));

  if (folded) {
    bands.push({
      label: "Other",
      values: periods.map((period) =>
        Object.entries(period.amounts ?? {})
          .filter(([label]) => !named.includes(label))
          .reduce((sum, [, amount]) => sum + (Number.isFinite(amount) && amount > 0 ? amount : 0), 0)
      ),
    });
  }

  const columnTotal = (index: number) =>
    bands.reduce((sum, band) => sum + (band.values[index] ?? 0), 0);

  const firstTotal = columnTotal(0);
  const lastTotal = columnTotal(periods.length - 1);

  // What moved, in share rather than in amount — which is the reading the
  // chart shows, and so the only one the sentence beside it may claim.
  const shareOf = (band: Band, index: number) => {
    const total = columnTotal(index);

    return total === 0 ? 0 : (band.values[index] ?? 0) / total;
  };

  const last = periods.length - 1;
  const share = (value: number) =>
    format.number(value, { style: "percent", maximumFractionDigits: 0 });

  const moved = bands
    .map((band) => ({
      label: band.label,
      from: shareOf(band, 0),
      to: shareOf(band, last),
    }))
    .sort((left, right) => Math.abs(right.to - right.from) - Math.abs(left.to - left.from))[0];

  const opening = periods[0].label;
  const closing = periods[last].label;

  return (
    <Section id="mix-over-time" title="How the mix has changed">
      <Paragraph id="mix-over-time-summary">
        {moved.to === moved.from
          ? `The mix has held steady from ${opening} to ${closing}.`
          : `${moved.label} ${moved.to > moved.from ? "grew" : "fell"} from ${
            share(moved.from)
          } of the total in ${opening} to ${share(moved.to)} in ${closing} — the largest move on ` +
            `the chart.`}
        {" "}
        The chart shows the mix rather than the size: the total itself went from{" "}
        {format.number(firstTotal)}{unit} to {format.number(lastTotal)}{unit}.
      </Paragraph>

      <Graph
        id="mix-over-time-chart"
        title="Share of each period"
        graphType="bar"
        // The whole point. `stacked` alone would draw the totals, and the
        // sentence above has just said the totals are not what this is for.
        stacked="percent"
        legend="bottom"
        data={{
          categories: periods.map((period) => period.label),
          // The raw figures, not the shares. Word divides.
          series: bands.map((band) => ({ label: band.label, values: band.values })),
        }}
        caption="Every column is drawn to the same height, so what is compared is the mix rather than the size."
      />
    </Section>
  );
};
