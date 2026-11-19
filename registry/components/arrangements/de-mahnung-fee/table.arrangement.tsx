import { Cell, Paragraph, Row, Table, type Yield } from "docxcelerate/template";
import type { DeMahnungFeeLines } from "./lines.ts";

/**
 * One row per invoice, so a letter covering many invoices can be scanned down
 * a column. A consumer has no Pauschale, so gets no column for it.
 */
export const table = (lines: DeMahnungFeeLines): Yield => (
  <>
    {lines.business
      ? (
        // 40mm holds an invoice heading in a few lines; the prose columns share the rest.
        <Table
          id="de-mahnung-fee-business-table"
          columns={[{ width: 40 }, { width: "auto" }, { width: "auto" }, { width: "auto" }]}
        >
          <Row id="de-mahnung-fee-business-header" header>
            <Cell>Rechnung</Cell>
            <Cell>Verzug</Cell>
            <Cell>Mahngebühr</Cell>
            <Cell>Verzugspauschale</Cell>
          </Row>
          {lines.invoices.map((invoice) => (
            <Row id="de-mahnung-fee-business-row">
              <Cell>{invoice.heading}</Cell>
              <Cell>{invoice.verzugLine}</Cell>
              <Cell>{invoice.feeLine}</Cell>
              <Cell>{invoice.pauschaleLine}</Cell>
            </Row>
          ))}
        </Table>
      )
      : (
        <Table id="de-mahnung-fee-consumer-table" columns={[{ width: 40 }, { width: "auto" }, { width: "auto" }]}>
          <Row id="de-mahnung-fee-consumer-header" header>
            <Cell>Rechnung</Cell>
            <Cell>Verzug</Cell>
            <Cell>Mahngebühr</Cell>
          </Row>
          {lines.invoices.map((invoice) => (
            <Row id="de-mahnung-fee-consumer-row">
              <Cell>{invoice.heading}</Cell>
              <Cell>{invoice.verzugLine}</Cell>
              <Cell>{invoice.feeLine}</Cell>
            </Row>
          ))}
        </Table>
      )}
    {lines.proofLine && <Paragraph id="de-mahnung-fee-proof">{lines.proofLine}</Paragraph>}
  </>
);
