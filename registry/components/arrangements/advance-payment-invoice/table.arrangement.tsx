import { Cell, Row, Table, type Yield } from "docxcelerate/template";
import type { AdvancePaymentInvoiceLines } from "./lines.ts";

/**
 * Labels in a column with the figures right-aligned beside them, so the
 * subtraction (previous plus this plus what is left is the order) can be
 * checked down the page. The default.
 */
export const table = (lines: AdvancePaymentInvoiceLines): Yield => (
  <>
    <Table id="advance-payment-figures" columns={[{ width: "auto" }, { width: 40, align: "right" }]}>
      <Row id="advance-payment-order-reference">
        <Cell>Order reference</Cell>
        <Cell>{lines.orderReference}</Cell>
      </Row>
      {lines.orderDate && (
        <Row id="advance-payment-order-date">
          <Cell>Order date</Cell>
          <Cell>{lines.orderDate}</Cell>
        </Row>
      )}
      <Row id="advance-payment-order-total">
        <Cell>Order total</Cell>
        <Cell>{lines.orderTotal}</Cell>
      </Row>
      {lines.previous.map((p) => (
        <Row id="advance-payment-previous">
          <Cell>Previously requested ({p.number})</Cell>
          <Cell>{p.amount}</Cell>
        </Row>
      ))}
      <Row id="advance-payment-deposit">
        <Cell>{lines.depositLabel}</Cell>
        <Cell>{lines.deposit}</Cell>
      </Row>
      <Row id="advance-payment-left">
        <Cell>Left to invoice</Cell>
        <Cell>{lines.leftToInvoice}</Cell>
      </Row>
      {lines.paymentDate && (
        <Row id="advance-payment-payment-date">
          <Cell>{lines.paymentDateLabel}</Cell>
          <Cell>{lines.paymentDate}</Cell>
        </Row>
      )}
    </Table>
    {lines.vatApplies && (
      <Table
        id="advance-payment-vat"
        columns={[{ width: "auto" }, { width: 34, align: "right" }, { width: 34, align: "right" }, {
          width: 34,
          align: "right",
        }]}
      >
        <Row header>
          <Cell>Rate</Cell>
          <Cell>Net</Cell>
          <Cell>VAT</Cell>
          <Cell>Gross</Cell>
        </Row>
        {lines.bands.map((band) => (
          <Row id="advance-payment-vat-band">
            <Cell>{band.label}</Cell>
            <Cell>{band.net}</Cell>
            <Cell>{band.vat}</Cell>
            <Cell>{band.gross}</Cell>
          </Row>
        ))}
        <Row id="advance-payment-vat-total">
          <Cell>Total</Cell>
          <Cell>{lines.netTotal}</Cell>
          <Cell>{lines.vatTotal}</Cell>
          <Cell>{lines.deposit}</Cell>
        </Row>
      </Table>
    )}
  </>
);
