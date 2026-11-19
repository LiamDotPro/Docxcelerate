import { Cell, Row, Table, type Yield } from "docxcelerate/template";
import type { NlCollectionCostsLines } from "./lines.ts";

/**
 * Labels on the left, amounts right-aligned, so the costs and btw visibly add
 * up to the total. The shape Dutch collection letters usually give the
 * hoofdsom, rente and kosten.
 */
export const table = (lines: NlCollectionCostsLines): Yield => (
  <Table id="nl-collection-costs-figures" columns={[{ width: "auto" }, { width: 35, align: "right" }]}>
    <Row id="nl-collection-costs-invoices">
      <Cell>Facturen</Cell>
      <Cell>{lines.invoices}</Cell>
    </Row>
    <Row id="nl-collection-costs-principal">
      <Cell>Hoofdsom</Cell>
      <Cell>{lines.principal}</Cell>
    </Row>
    {lines.interest && (
      <Row id="nl-collection-costs-interest">
        <Cell>Wettelijke rente bijgeteld</Cell>
        <Cell>{lines.interest}</Cell>
      </Row>
    )}
    <Row id="nl-collection-costs-costs">
      <Cell>Incassokosten</Cell>
      <Cell>{lines.costs}</Cell>
    </Row>
    {lines.vat && (
      <Row id="nl-collection-costs-vat">
        <Cell>Btw (21%) over de incassokosten</Cell>
        <Cell>{lines.vat}</Cell>
      </Row>
    )}
    <Row id="nl-collection-costs-total">
      <Cell>{lines.totalLabel}</Cell>
      <Cell>{lines.total}</Cell>
    </Row>
  </Table>
);
