import {
  Graph,
  Paragraph,
  Section,
  type Section as SectionComponent,
  useFormat,
  useState,
} from "docxcelerate/template";

/**
 * Where a total went, as a pie and as the sentence a reader actually reads.
 *
 * The decision this component exists to make for you is **what happens past
 * the sixth slice**. A breakdown arriving from real data has as many
 * categories as the system had rows, and a chart drawn straight from it runs
 * off the end of the palette — where the ninth slice is painted the same as
 * the first, and two unrelated things look like one. So everything past the
 * largest few is added up and drawn as one slice called "Other", which is what
 * a person reading a pie can hold in their head anyway.
 *
 * Two more decisions come with it.
 *
 * **The slices are sorted largest first**, because a pie read clockwise from
 * the top is read in order, and an unsorted one asks the reader to do the
 * ranking themselves.
 *
 * **Every slice prints its share**, so identity is never left to colour alone.
 * Three of the shipped palette's hues sit under 3:1 against white; a reader
 * who cannot tell two of them apart still has the number and the key. It is a
 * pie rather than a doughnut for the same reason: measured in Word, a pie puts
 * the labels for its smallest slices outside the circle with a leader line,
 * and a doughnut crowds them into the ring where the 2% and the 4% overlap.
 *
 * Installed by `dxcl add usage-breakdown`.
 */

/**
 * How many slices are drawn before the rest become "Other".
 *
 * Six, because that is where the shipped chart palette stops being
 * comfortably distinguishable and well before it stops being distinguishable
 * at all. Raise it and check the result on a printed page, not a screen.
 */
const SLICES = 6;

/** What this component reads. Add these fields to your document data type. */
export interface UsageBreakdownData {
  usage: {
    /**
     * One entry per category, in any order.
     *
     * Order does not matter here — unlike a trend, a breakdown has no
     * sequence of its own, so this component sorts it.
     */
    items: Array<{
      /** What the category is called, as it should be printed. */
      label: string;
      /** How much of the total it accounts for. Negatives are dropped. */
      amount: number;
    }>;
    /** What is being counted: `"kWh"`, `"visits"`, `"hours"`. Printed in the prose. */
    unit?: string;
    /** What the breakdown covers: `"this quarter"`, `"since April"`. */
    period?: string;
  };
}

export const UsageBreakdown: SectionComponent = () => {
  const format = useFormat();
  const [usage] = useState((data: UsageBreakdownData) => data.usage);

  // A slice of nothing is not a slice, and a negative one cannot be drawn as
  // an angle at all — a refund in a breakdown of where the money went belongs
  // in the prose, not in the ring.
  const items = (usage.items ?? [])
    .filter((entry) => Number.isFinite(entry.amount) && entry.amount > 0)
    .sort((left, right) => right.amount - left.amount);

  const total = items.reduce((sum, entry) => sum + entry.amount, 0);
  const unit = usage.unit === undefined ? "" : ` ${usage.unit}`;
  const period = usage.period === undefined ? "" : ` ${usage.period}`;

  if (items.length === 0) {
    return (
      <Section id="usage-breakdown" title="Where it went">
        <Paragraph id="usage-breakdown-none">
          We have nothing recorded against this account{period}.
        </Paragraph>
      </Section>
    );
  }

  // Everything past the largest few, added up. `slice` is safe on a short
  // array, so a breakdown of three categories simply never grows an "Other".
  const shown = items.slice(0, SLICES);
  const rest = items.slice(SLICES);
  const other = rest.reduce((sum, entry) => sum + entry.amount, 0);

  const slices = other > 0 ? [...shown, { label: "Other", amount: other }] : shown;
  const largest = slices[0];
  const share = (amount: number) =>
    format.number(amount / total, { style: "percent", maximumFractionDigits: 0 });

  return (
    <Section id="usage-breakdown" title="Where it went">
      <Paragraph id="usage-breakdown-summary">
        You used {format.number(total)}{unit}{period}.{" "}
        {largest.label} accounts for {share(largest.amount)} of it
        {rest.length > 0
          ? `, and the ${format.number(rest.length)} smallest ` +
            `${format.plural(rest.length, "category", "categories")} are shown together as Other.`
          : "."}
      </Paragraph>

      <Graph
        id="usage-breakdown-chart"
        title="Share of the total"
        graphType="pie"
        // The key names the slices rather than the series, so it stays even
        // though there is only one run of numbers here.
        legend="bottom"
        // The share on every slice, because three of the palette's hues sit
        // under 3:1 on white and colour alone is not an answer.
        dataLabels
        data={{
          categories: slices.map((entry) => entry.label),
          series: [{ label: "Share", values: slices.map((entry) => entry.amount) }],
        }}
        caption={`How the ${format.number(total)}${unit} breaks down${period}.`}
      />
    </Section>
  );
};
