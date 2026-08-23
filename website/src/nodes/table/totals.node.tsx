import { Cell, Paragraph, Row, Table, useFormat, useState } from "docxcelerate/template";
import type { SampleData } from "../sample-data.ts";

/**
 * A cell takes text directly, and paragraphs when one line is not enough. The
 * closing row is marked `header` so it is drawn as one — it stays where it is,
 * because only the rows a table opens with repeat onto a new page.
 */
export const PriceSummary: Table = () => {
  const { currency } = useFormat("en-GB");
  const [state] = useState((data: SampleData) => ({
    plan: data.plan,
    lastPrice: data.lastPrice,
    newPrice: data.newPrice,
  }));

  return (
    <Table id="price-summary" columns={[{ width: "auto" }, { width: 30, align: "right" }]}>
      <Row>
        <Cell>
          <Paragraph>{state.plan}</Paragraph>
          <Paragraph>Last year</Paragraph>
        </Cell>
        <Cell>{currency(state.lastPrice)}</Cell>
      </Row>
      <Row header>
        <Cell>From renewal</Cell>
        <Cell>{currency(state.newPrice)}</Cell>
      </Row>
    </Table>
  );
};
