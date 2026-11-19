import { Cell, Paragraph, Row, Table, type Yield } from "docxcelerate/template";
import type { JapanQualifiedInvoiceLines } from "./lines.ts";

/**
 * Items and band totals in one two-column table, amounts lined up on the
 * right, so the per-rate totals can be checked by eye. The conventional
 * business invoice.
 */
export const table = (lines: JapanQualifiedInvoiceLines): Yield => (
  <>
    {lines.buyer && <Paragraph id="jp-invoice-buyer">{lines.buyer}</Paragraph>}
    <Paragraph id="jp-invoice-date">取引年月日：{lines.date}</Paragraph>
    {lines.registration && <Paragraph id="jp-invoice-registration">登録番号：{lines.registration}</Paragraph>}
    <Table id="jp-invoice-lines" columns={[{ width: "auto" }, { width: 40, align: "right" }]}>
      <Row id="jp-invoice-lines-header" header>
        <Cell>品名</Cell>
        <Cell>{lines.amountHeader}</Cell>
      </Row>
      {lines.items[0].description && (
        <Row id="jp-invoice-line-1">
          <Cell>{lines.items[0].description}</Cell>
          <Cell>{lines.items[0].amount}</Cell>
        </Row>
      )}
      {lines.items[1].description && (
        <Row id="jp-invoice-line-2">
          <Cell>{lines.items[1].description}</Cell>
          <Cell>{lines.items[1].amount}</Cell>
        </Row>
      )}
      {lines.items[2].description && (
        <Row id="jp-invoice-line-3">
          <Cell>{lines.items[2].description}</Cell>
          <Cell>{lines.items[2].amount}</Cell>
        </Row>
      )}
      {lines.items[3].description && (
        <Row id="jp-invoice-line-4">
          <Cell>{lines.items[3].description}</Cell>
          <Cell>{lines.items[3].amount}</Cell>
        </Row>
      )}
      {lines.items[4].description && (
        <Row id="jp-invoice-line-5">
          <Cell>{lines.items[4].description}</Cell>
          <Cell>{lines.items[4].amount}</Cell>
        </Row>
      )}
      {lines.items[5].description && (
        <Row id="jp-invoice-line-6">
          <Cell>{lines.items[5].description}</Cell>
          <Cell>{lines.items[5].amount}</Cell>
        </Row>
      )}
      {lines.items[6].description && (
        <Row id="jp-invoice-line-7">
          <Cell>{lines.items[6].description}</Cell>
          <Cell>{lines.items[6].amount}</Cell>
        </Row>
      )}
      {lines.items[7].description && (
        <Row id="jp-invoice-line-8">
          <Cell>{lines.items[7].description}</Cell>
          <Cell>{lines.items[7].amount}</Cell>
        </Row>
      )}
      {lines.items[8].description && (
        <Row id="jp-invoice-line-9">
          <Cell>{lines.items[8].description}</Cell>
          <Cell>{lines.items[8].amount}</Cell>
        </Row>
      )}
      {lines.items[9].description && (
        <Row id="jp-invoice-line-10">
          <Cell>{lines.items[9].description}</Cell>
          <Cell>{lines.items[9].amount}</Cell>
        </Row>
      )}
      {lines.standard.total && (
        <Row id="jp-invoice-standard-total">
          <Cell>{lines.standard.label}</Cell>
          <Cell>{lines.standard.total}</Cell>
        </Row>
      )}
      {lines.standard.tax && (
        <Row id="jp-invoice-standard-tax">
          <Cell>{lines.standard.taxLabel}</Cell>
          <Cell>{lines.standard.tax}</Cell>
        </Row>
      )}
      {lines.reduced.total && (
        <Row id="jp-invoice-reduced-total">
          <Cell>{lines.reduced.label}</Cell>
          <Cell>{lines.reduced.total}</Cell>
        </Row>
      )}
      {lines.reduced.tax && (
        <Row id="jp-invoice-reduced-tax">
          <Cell>{lines.reduced.taxLabel}</Cell>
          <Cell>{lines.reduced.tax}</Cell>
        </Row>
      )}
      <Row id="jp-invoice-total">
        <Cell>{lines.totalLabel}</Cell>
        <Cell>{lines.total}</Cell>
      </Row>
    </Table>
    {lines.legend && <Paragraph id="jp-invoice-legend">{lines.legend}</Paragraph>}
    {lines.note && <Paragraph id="jp-invoice-note">{lines.note}</Paragraph>}
  </>
);
