import { Cell, Paragraph, type Row as RowComponent, Row, useDeriver } from "docxcelerate/template";
import { chargeLine } from "../derivers.ts";
import type { InvoiceLine } from "../types.ts";

/**
 * One line of the charges table.
 *
 * A row of its own component rather than four cells written inline, because
 * the figures have to be derived and a deriver is a hook: hooks belong to a
 * component, and the callback inside a `.map()` is not one. Published, this
 * becomes the body of the loop the engine walks — one row written once,
 * standing for however many the request turns out to have.
 */
export const ChargeRow: RowComponent<{ line: InvoiceLine }> = async ({ line }) => {
  const figures = await useDeriver(chargeLine, [line.qty, line.rate]);

  return (
    <Row>
      <Cell variant="lineItem">
        <Paragraph>{line.desc}</Paragraph>
        <Paragraph variant="chargeNote">{line.meta}</Paragraph>
      </Cell>
      <Cell variant="money">{figures.qty}</Cell>
      <Cell variant="money">{figures.rate}</Cell>
      <Cell variant="money">{figures.amount}</Cell>
    </Row>
  );
};
