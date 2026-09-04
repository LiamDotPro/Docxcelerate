import {
  Graph,
  Paragraph,
  Section,
  type Section as SectionComponent,
  useFormat,
  useState,
} from "docxcelerate/template";

/**
 * How a balance has moved, as a line and as a sentence saying the same thing.
 *
 * The two are worked out from one array, which is the whole reason to install
 * this rather than write a `<Graph>` yourself: a chart that says the balance
 * fell beside a sentence that says it rose is the failure this shape rules
 * out. Everything printed comes from `history`, and nothing is passed in twice.
 *
 * Three decisions are already made here.
 *
 * **A reading nobody took is `null`, not `0`.** A month that has not been
 * billed yet is not a month that was billed nothing, and drawing it as zero
 * puts a cliff in the line that never happened. The chart draws a gap, and the
 * sentence measures from the first reading it has to the last.
 *
 * **Fewer than two readings is not a chart.** One point is not a trend, and a
 * line drawn through it is a chart that looks like it says something. The
 * component prints the balance as prose instead and draws nothing.
 *
 * **The prose carries the currency and the axis carries the number.** A value
 * axis repeating "£" down its side is six copies of a fact the sentence
 * already gave. Set `numberFormat` if you want it there too.
 *
 * Installed by `dxcl add balance-trend`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface BalanceTrendData {
  balance: {
    /**
     * One entry per period, oldest first.
     *
     * The order is the chart's order. Sort before it reaches the document —
     * a chart is drawn in the order it is given, and a component that sorted
     * for you would disagree with the table beside it.
     */
    history: Array<{
      /** What the period is called: `"Jan"`, `"Q1"`, `"Week 12"`. */
      period: string;
      /** What was outstanding at the end of it. `null` where nothing was read. */
      amount: number | null;
    }>;
    /** ISO 4217 code, for the sentence. Defaults to GBP. */
    currency?: string;
    /**
     * How the value axis prints its figures, as an OOXML number format.
     *
     * Defaults to `"#,##0"` — grouped whole numbers. `"£#,##0"` puts the
     * symbol on the axis as well as in the sentence.
     */
    numberFormat?: string;
  };
}

export const BalanceTrend: SectionComponent = () => {
  const format = useFormat();
  const [balance] = useState((data: BalanceTrendData) => data.balance);

  const history = balance.history ?? [];
  const currency = balance.currency ?? "GBP";
  // The readings that exist, with where each sits kept, so the sentence talks
  // about the same periods the chart plots.
  const read = history
    .map((entry, index) => ({ ...entry, index }))
    .filter((entry): entry is { period: string; amount: number; index: number } =>
      typeof entry.amount === "number"
    );

  const first = read[0];
  const last = read[read.length - 1];
  const money = (amount: number) => format.currency(amount, currency);

  if (last === undefined) {
    return (
      <Section id="balance-trend" title="How your balance has moved">
        <Paragraph id="balance-trend-none">
          We have no readings for this account yet, so there is nothing to show.
        </Paragraph>
      </Section>
    );
  }

  // One reading is a balance, not a trend. Said as a balance.
  if (read.length < 2) {
    return (
      <Section id="balance-trend" title="Your balance">
        <Paragraph id="balance-trend-single">
          The balance on your account is {money(last.amount)}, as at {last.period}.
        </Paragraph>
      </Section>
    );
  }

  const change = last.amount - first.amount;
  const moved = Math.abs(change) >= 0.005;

  return (
    <Section id="balance-trend" title="How your balance has moved">
      <Paragraph id="balance-trend-summary">
        {moved
          ? `Between ${first.period} and ${last.period} your balance ` +
            `${change > 0 ? "rose" : "fell"} by ${money(Math.abs(change))}, ` +
            `from ${money(first.amount)} to ${money(last.amount)}.`
          : `Your balance has stayed at ${money(last.amount)} ` +
            `from ${first.period} to ${last.period}.`}
      </Paragraph>

      <Graph
        id="balance-trend-chart"
        title="Balance by period"
        graphType="line"
        // One series needs no key: it would name the only thing on the chart,
        // which the title has already named.
        legend="none"
        valueAxisTitle="Balance"
        numberFormat={balance.numberFormat ?? "#,##0"}
        data={{
          categories: history.map((entry) => entry.period),
          series: [{ label: "Balance", values: history.map((entry) => entry.amount ?? null) }],
        }}
        caption={`Balance at the end of each period, ${first.period} to ${last.period}.`}
      />
    </Section>
  );
};
