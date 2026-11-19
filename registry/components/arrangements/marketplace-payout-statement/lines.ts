/**
 * One row of the statement, already decided. An arrangement places it and
 * never works anything out.
 *
 * Installed by `dxcl add marketplace-payout-statement`, with the arrangements.
 */
export interface PayoutLine {
  /** The row's label, as printed. */
  label: string;
  /** The signed amount ("+£12.00", "-£3.50"), or "" when the row is not printed. */
  amount: string;
  /** The unsigned amount when the row adds money (or is zero), else "". */
  paidIn: string;
  /** The unsigned amount when the row takes money away, else "". */
  takenOut: string;
}

/**
 * Every row the statement may print, in reconciliation order. Test each one in
 * the JSX, `{lines.refunds.amount && …}`, so the test is compiled into a
 * condition the engine settles per document.
 */
export interface MarketplacePayoutStatementLines {
  /** The balance the last statement could not pay out. */
  broughtForward: PayoutLine;
  /** Items and shipping sold. Always printed. */
  sales: PayoutLine;
  /** Facilitator sales tax collected from buyers. */
  taxCollected: PayoutLine;
  /** The same tax, taken straight back out. */
  taxRemitted: PayoutLine;
  /** What went back to buyers. */
  refunds: PayoutLine;
  /** Commission credited back on those refunds, less the administration fee. */
  commissionReturned: PayoutLine;
  /** The administration fee the marketplace keeps on refunds. */
  refundAdminFees: PayoutLine;
  /** Commission and every other fee, before VAT. */
  fees: PayoutLine;
  /** VAT on the fees. */
  feeVat: PayoutLine;
  /** Chargebacks raised this period. */
  chargebacks: PayoutLine;
  /** Dispute fees on those chargebacks, kept even on a win. */
  disputeFees: PayoutLine;
  /** Earlier deductions reversed by a won dispute. */
  disputesWon: PayoutLine;
  /** New reserve holds. */
  reserveHeld: PayoutLine;
  /** Reserve holds that matured. */
  reserveReleased: PayoutLine;
  /** The closing balance. Always printed. */
  closing: PayoutLine;
}
