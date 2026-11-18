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
      {lines.previous[0].number && (
        <Row id="advance-payment-previous-1">
          <Cell>Previously requested ({lines.previous[0].number})</Cell>
          <Cell>{lines.previous[0].amount}</Cell>
        </Row>
      )}
      {lines.previous[1].number && (
        <Row id="advance-payment-previous-2">
          <Cell>Previously requested ({lines.previous[1].number})</Cell>
          <Cell>{lines.previous[1].amount}</Cell>
        </Row>
      )}
      {lines.previous[2].number && (
        <Row id="advance-payment-previous-3">
          <Cell>Previously requested ({lines.previous[2].number})</Cell>
          <Cell>{lines.previous[2].amount}</Cell>
        </Row>
      )}
      {lines.previous[3].number && (
        <Row id="advance-payment-previous-4">
          <Cell>Previously requested ({lines.previous[3].number})</Cell>
          <Cell>{lines.previous[3].amount}</Cell>
        </Row>
      )}
      {lines.previous[4].number && (
        <Row id="advance-payment-previous-5">
          <Cell>Previously requested ({lines.previous[4].number})</Cell>
          <Cell>{lines.previous[4].amount}</Cell>
        </Row>
      )}
      {lines.previous[5].number && (
        <Row id="advance-payment-previous-6">
          <Cell>Previously requested ({lines.previous[5].number})</Cell>
          <Cell>{lines.previous[5].amount}</Cell>
        </Row>
      )}
      {lines.previous[6].number && (
        <Row id="advance-payment-previous-7">
          <Cell>Previously requested ({lines.previous[6].number})</Cell>
          <Cell>{lines.previous[6].amount}</Cell>
        </Row>
      )}
      {lines.previous[7].number && (
        <Row id="advance-payment-previous-8">
          <Cell>Previously requested ({lines.previous[7].number})</Cell>
          <Cell>{lines.previous[7].amount}</Cell>
        </Row>
      )}
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
        {lines.bands[0].label && (
          <Row id="advance-payment-vat-band-1">
            <Cell>{lines.bands[0].label}</Cell>
            <Cell>{lines.bands[0].net}</Cell>
            <Cell>{lines.bands[0].vat}</Cell>
            <Cell>{lines.bands[0].gross}</Cell>
          </Row>
        )}
        {lines.bands[1].label && (
          <Row id="advance-payment-vat-band-2">
            <Cell>{lines.bands[1].label}</Cell>
            <Cell>{lines.bands[1].net}</Cell>
            <Cell>{lines.bands[1].vat}</Cell>
            <Cell>{lines.bands[1].gross}</Cell>
          </Row>
        )}
        {lines.bands[2].label && (
          <Row id="advance-payment-vat-band-3">
            <Cell>{lines.bands[2].label}</Cell>
            <Cell>{lines.bands[2].net}</Cell>
            <Cell>{lines.bands[2].vat}</Cell>
            <Cell>{lines.bands[2].gross}</Cell>
          </Row>
        )}
        {lines.bands[3].label && (
          <Row id="advance-payment-vat-band-4">
            <Cell>{lines.bands[3].label}</Cell>
            <Cell>{lines.bands[3].net}</Cell>
            <Cell>{lines.bands[3].vat}</Cell>
            <Cell>{lines.bands[3].gross}</Cell>
          </Row>
        )}
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
