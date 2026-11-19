/**
 * What the deposit outcome hands each of its arrangements: every line it may
 * print, already formatted and already decided. An arrangement places these
 * and decides nothing, so both print the same lines under the same conditions.
 *
 * A `null` is a line this order does not carry. Test it in the JSX,
 * `{lines.security && …}`, so the test is compiled into a condition the engine
 * settles per document.
 *
 * Installed by `dxcl add deposit-cancellation-outcome`, with the arrangements.
 */
export interface DepositCancellationOutcomeLines {
  /** Total paid in advance, formatted; null when nothing was paid (the arrangement prints nothing). */
  paid: string | null;
  /** Security deposit held; null when none. */
  security: string | null;
  /** Exactly 6 slots. An unused slot has every field null. */
  deductions: Array<{
    /** "Specialist cleaning (including £308.40 VAT)"; US: the label alone. */
    description: string | null;
    /** The deduction, formatted; null on an unused slot. */
    amount: string | null;
    /** "Specialist cleaning: £1,850.40, including £308.40 VAT", for stacked. */
    line: string | null;
  }>;
  /** Pooled refundable advance paid; null when none. */
  refundable: string | null;
  /** "Retained: <reason>"; null when nothing is retained from the refundable advances. */
  retentionLabel: string | null;
  /** The amount retained from the refundable advances; null when zero. */
  retention: string | null;
  /** Non-refundable advance, retained in full; null when none. */
  nonRefundable: string | null;
  /** Refunded to you. Non-null whenever anything was paid, even "£0.00". */
  refunded: string | null;
  /** Total retained (deductions + retention + non-refundable); null when zero. */
  retained: string | null;
}
