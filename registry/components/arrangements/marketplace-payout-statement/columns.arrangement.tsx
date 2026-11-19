import { Cell, Row, Table, type Yield } from "docxcelerate/template";
import type { MarketplacePayoutStatementLines } from "./lines.ts";

/**
 * Money in and money out in two unsigned columns, the way a bank statement
 * reads. The closing balance goes in "Paid in" unless it is negative.
 */
export const columns = (lines: MarketplacePayoutStatementLines): Yield => (
  <Table id="payout-lines" columns={[{ width: "auto" }, { width: 35, align: "right" }, { width: 35, align: "right" }]}>
    <Row id="payout-lines-header" header>
      <Cell>Description</Cell>
      <Cell>Paid in</Cell>
      <Cell>Taken out</Cell>
    </Row>
    {lines.broughtForward.amount && (
      <Row id="payout-brought-forward">
        <Cell>{lines.broughtForward.label}</Cell>
        <Cell>{lines.broughtForward.paidIn}</Cell>
        <Cell>{lines.broughtForward.takenOut}</Cell>
      </Row>
    )}
    <Row id="payout-sales">
      <Cell>{lines.sales.label}</Cell>
      <Cell>{lines.sales.paidIn}</Cell>
      <Cell>{lines.sales.takenOut}</Cell>
    </Row>
    {lines.taxCollected.amount && (
      <Row id="payout-tax-collected">
        <Cell>{lines.taxCollected.label}</Cell>
        <Cell>{lines.taxCollected.paidIn}</Cell>
        <Cell>{lines.taxCollected.takenOut}</Cell>
      </Row>
    )}
    {lines.taxCollected.amount && (
      <Row id="payout-tax-remitted">
        <Cell>{lines.taxRemitted.label}</Cell>
        <Cell>{lines.taxRemitted.paidIn}</Cell>
        <Cell>{lines.taxRemitted.takenOut}</Cell>
      </Row>
    )}
    {lines.refunds.amount && (
      <Row id="payout-refunds">
        <Cell>{lines.refunds.label}</Cell>
        <Cell>{lines.refunds.paidIn}</Cell>
        <Cell>{lines.refunds.takenOut}</Cell>
      </Row>
    )}
    {lines.commissionReturned.amount && (
      <Row id="payout-commission-returned">
        <Cell>{lines.commissionReturned.label}</Cell>
        <Cell>{lines.commissionReturned.paidIn}</Cell>
        <Cell>{lines.commissionReturned.takenOut}</Cell>
      </Row>
    )}
    {lines.refundAdminFees.amount && (
      <Row id="payout-refund-admin-fees">
        <Cell>{lines.refundAdminFees.label}</Cell>
        <Cell>{lines.refundAdminFees.paidIn}</Cell>
        <Cell>{lines.refundAdminFees.takenOut}</Cell>
      </Row>
    )}
    {lines.fees.amount && (
      <Row id="payout-fees">
        <Cell>{lines.fees.label}</Cell>
        <Cell>{lines.fees.paidIn}</Cell>
        <Cell>{lines.fees.takenOut}</Cell>
      </Row>
    )}
    {lines.feeVat.amount && (
      <Row id="payout-fee-vat">
        <Cell>{lines.feeVat.label}</Cell>
        <Cell>{lines.feeVat.paidIn}</Cell>
        <Cell>{lines.feeVat.takenOut}</Cell>
      </Row>
    )}
    {lines.chargebacks.amount && (
      <Row id="payout-chargebacks">
        <Cell>{lines.chargebacks.label}</Cell>
        <Cell>{lines.chargebacks.paidIn}</Cell>
        <Cell>{lines.chargebacks.takenOut}</Cell>
      </Row>
    )}
    {lines.disputeFees.amount && (
      <Row id="payout-dispute-fees">
        <Cell>{lines.disputeFees.label}</Cell>
        <Cell>{lines.disputeFees.paidIn}</Cell>
        <Cell>{lines.disputeFees.takenOut}</Cell>
      </Row>
    )}
    {lines.disputesWon.amount && (
      <Row id="payout-disputes-won">
        <Cell>{lines.disputesWon.label}</Cell>
        <Cell>{lines.disputesWon.paidIn}</Cell>
        <Cell>{lines.disputesWon.takenOut}</Cell>
      </Row>
    )}
    {lines.reserveHeld.amount && (
      <Row id="payout-reserve-held">
        <Cell>{lines.reserveHeld.label}</Cell>
        <Cell>{lines.reserveHeld.paidIn}</Cell>
        <Cell>{lines.reserveHeld.takenOut}</Cell>
      </Row>
    )}
    {lines.reserveReleased.amount && (
      <Row id="payout-reserve-released">
        <Cell>{lines.reserveReleased.label}</Cell>
        <Cell>{lines.reserveReleased.paidIn}</Cell>
        <Cell>{lines.reserveReleased.takenOut}</Cell>
      </Row>
    )}
    <Row id="payout-closing">
      <Cell>{lines.closing.label}</Cell>
      <Cell>{lines.closing.paidIn}</Cell>
      <Cell>{lines.closing.takenOut}</Cell>
    </Row>
  </Table>
);
