import { Cell, type Nodes, Paragraph, Row, Table, Text, useState } from "docxcelerate/template";
import { percent } from "../format.ts";
import { Bar, BarList, Opening, Rubric } from "./parts.tsx";
import { Written } from "./written.node.tsx";
import type { PortfolioReviewData } from "../types.ts";

/** 178mm less the 18mm label, the 22mm and 20mm figures and three 8pt gaps. */
const DRIFT_TRACK = 109.54;
/** How many millimetres of track one percentage point takes: 80% of it holds 64%. */
const PER_POINT = DRIFT_TRACK / 80;
const GUTTER = 2.82;


/** Page 6: the mix against the agreed one, then where the money is by region and currency. */
export const Allocation: Nodes = () => {
  const [mix] = useState((data: PortfolioReviewData) => ({
    classes: data.allocation,
    regions: data.regions,
    currencies: data.currencies,
  }));

  const largest = Math.max(...mix.regions.map((region) => region.percent));
  const regionTrack = 178 - 40 - 12 - 2 * GUTTER;
  const total = mix.currencies.reduce((sum, currency) => sum + currency.percent, 0);
  // Built before the paragraph rather than mapped inside it: a map in JSX is a
  // loop the engine walks, and a paragraph holds runs, not loops.
  const key = mix.currencies.flatMap((currency, index) => [
    <Text variant={`key${index}`}>■</Text>,
    ` ${currency.name} ${currency.percent}%        `,
  ]);

  return (
    <>
      <Opening id="allocation" number="4" title="Asset allocation" standfirst="Your mix against the one we agreed, and where the money is invested." />

      <Rubric id="agreed-mix" label="Against your agreed mix" note="■ Now   ▏Target   ■ Agreed range" variant="rubricTight" />
      <Table
        id="drift-table"
        variant="drift"
        columns={[
          { width: 18 }, { width: GUTTER }, { width: DRIFT_TRACK }, { width: GUTTER },
          { width: 22, align: "right" }, { width: GUTTER }, { width: 20, align: "right" },
        ]}
      >
        <Row variant="driftHead">
          <Cell></Cell>
          <Cell></Cell>
          <Cell></Cell>
          <Cell></Cell>
          <Cell>Now / target</Cell>
          <Cell></Cell>
          <Cell>After trades</Cell>
        </Row>
        {mix.classes.map((row, index) => {
          const outside = Math.abs(row.now - row.target) > row.tolerance;
          return (
            <>
              <Row height={index === 0 ? 3 : 11} heightRule="exact">
                {[1, 2, 3, 4, 5, 6, 7].map(() => <Cell variant="barNone"></Cell>)}
              </Row>
              <Row>
                <Cell variant="driftName">{row.name}</Cell>
                <Cell></Cell>
                <Cell>
                  <Bar
                    id={`drift-${row.name.toLowerCase()}`}
                    width={DRIFT_TRACK}
                    heightPt={12}
                    insetPt={3}
                    trailingPt={1}
                    track="barTrack"
                    marks={[
                      { from: (row.target - row.tolerance) * PER_POINT, to: (row.target + row.tolerance) * PER_POINT, variant: "barRange" },
                      { from: 0, to: row.now * PER_POINT, variant: outside ? "barDeep" : "barMid", inset: true },
                      { from: row.target * PER_POINT - 0.26, to: row.target * PER_POINT + 0.27, variant: "barAxis" },
                    ]}
                  />
                </Cell>
                <Cell></Cell>
                <Cell>
                  <Text variant="driftNow">{percent(row.now)}</Text>
                  <Text variant="driftTarget">{` / ${row.target}%`}</Text>
                </Cell>
                <Cell></Cell>
                <Cell variant="driftAfter">{percent(row.after)}</Cell>
              </Row>
            </>
          );
        })}
      </Table>
      <Written
        id="drift"
        variant="caption"
        ask="Two sentences on how far each class has drifted from target and what the proposed trades do about it."
        placeholder="What the drift from the agreed mix means for this client."
        from={{ allocation: mix.classes }}
      />

      <Rubric id="regions" label="By region" note="Looking through to what each fund holds" />
      <BarList
        id="region-bars"
        labelWidth={40}
        valueWidth={12}
        heightPt={12}
        gapPt={7}
        lines={mix.regions.map((region) => ({
          label: region.name,
          value: `${region.percent}%`,
          valueVariant: "barValuePlain",
          marks: [{ from: 0, to: (region.percent / largest) * regionTrack, variant: "barGreen" }],
        }))}
      />

      <Rubric id="currencies" label="By currency" note="Bond holdings are hedged to sterling" />
      <Table id="currency-bar" variant="currency" columns={mix.currencies.map((currency) => ({ width: (currency.percent / total) * 178, align: "center" as const }))}>
        <Row height={22} heightRule="exact">
          {mix.currencies.map((currency, index) => (
            <Cell variant={`currency${index}`}>
              {currency.code ? `${currency.code} ${currency.percent}%` : ""}
            </Cell>
          ))}
        </Row>
      </Table>
      <Paragraph id="currency-key" variant="currencyKey">{key}</Paragraph>
    </>
  );
};
