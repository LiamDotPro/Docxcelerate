import { Cell, Row, Table, type Yield } from "docxcelerate/template";
import type { MarketplacePayoutStatementLines } from "./lines.ts";

/**
 * One amount column with every row signed, so a reader adds the column down to
 * the closing balance and sees at a glance which rows took money away.
 */
export const signed = (lines: MarketplacePayoutStatementLines): Yield => (
  <Table id="payout-lines" columns={[{ width: "auto" }, { width: 40, align: "right" }]}>
    <Row id="payout-lines-header" header>
      <Cell>Description</Cell>
      <Cell>Amount</Cell>
    </Row>
    {lines.broughtForward.amount && (
      <Row id="payout-brought-forward">
        <Cell>{lines.broughtForward.label}</Cell>
        <Cell>{lines.broughtForward.amount}</Cell>
      </Row>
    )}
    <Row id="payout-sales">
      <Cell>{lines.sales.label}</Cell>
      <Cell>{lines.sales.amount}</Cell>
    </Row>
    {lines.taxCollected.amount && (
      <Row id="payout-tax-collected">
        <Cell>{lines.taxCollected.label}</Cell>
        <Cell>{lines.taxCollected.amount}</Cell>
      </Row>
    )}
    {lines.taxCollected.amount && (
      <Row id="payout-tax-remitted">
        <Cell>{lines.taxRemitted.label}</Cell>
        <Cell>{lines.taxRemitted.amount}</Cell>
      </Row>
    )}
    {lines.refunds.amount && (
      <Row id="payout-refunds">
        <Cell>{lines.refunds.label}</Cell>
        <Cell>{lines.refunds.amount}</Cell>
      </Row>
    )}
    {lines.commissionReturned.amount && (
      <Row id="payout-commission-returned">
        <Cell>{lines.commissionReturned.label}</Cell>
        <Cell>{lines.commissionReturned.amount}</Cell>
      </Row>
    )}
    {lines.refundAdminFees.amount && (
      <Row id="payout-refund-admin-fees">
        <Cell>{lines.refundAdminFees.label}</Cell>
        <Cell>{lines.refundAdminFees.amount}</Cell>
      </Row>
    )}
    {lines.fees.amount && (
      <Row id="payout-fees">
        <Cell>{lines.fees.label}</Cell>
        <Cell>{lines.fees.amount}</Cell>
      </Row>
    )}
    {lines.feeVat.amount && (
      <Row id="payout-fee-vat">
        <Cell>{lines.feeVat.label}</Cell>
        <Cell>{lines.feeVat.amount}</Cell>
      </Row>
    )}
    {lines.chargebacks.amount && (
      <Row id="payout-chargebacks">
        <Cell>{lines.chargebacks.label}</Cell>
        <Cell>{lines.chargebacks.amount}</Cell>
      </Row>
    )}
    {lines.disputeFees.amount && (
      <Row id="payout-dispute-fees">
        <Cell>{lines.disputeFees.label}</Cell>
        <Cell>{lines.disputeFees.amount}</Cell>
      </Row>
    )}
    {lines.disputesWon.amount && (
      <Row id="payout-disputes-won">
        <Cell>{lines.disputesWon.label}</Cell>
        <Cell>{lines.disputesWon.amount}</Cell>
      </Row>
    )}
    {lines.reserveHeld.amount && (
      <Row id="payout-reserve-held">
        <Cell>{lines.reserveHeld.label}</Cell>
        <Cell>{lines.reserveHeld.amount}</Cell>
      </Row>
    )}
    {lines.reserveReleased.amount && (
      <Row id="payout-reserve-released">
        <Cell>{lines.reserveReleased.label}</Cell>
        <Cell>{lines.reserveReleased.amount}</Cell>
      </Row>
    )}
    <Row id="payout-closing">
      <Cell>{lines.closing.label}</Cell>
      <Cell>{lines.closing.amount}</Cell>
    </Row>
  </Table>
);
