import { Paragraph, type Yield } from "docxcelerate/template";
import type { DepositCancellationOutcomeLines } from "./lines.ts";

/** One "Label: amount" line under another, down the text column. */
export const stacked = (lines: DepositCancellationOutcomeLines): Yield => (
  <>
    {lines.paid && (
      <>
      <Paragraph id="deposit-cancellation-paid">Paid in advance: {lines.paid}</Paragraph>
      {lines.security && <Paragraph id="deposit-cancellation-security">Security deposit: {lines.security}</Paragraph>}
      {lines.deductions[0].line && <Paragraph id="deposit-cancellation-deduction-1">{lines.deductions[0].line}</Paragraph>}
      {lines.deductions[1].line && <Paragraph id="deposit-cancellation-deduction-2">{lines.deductions[1].line}</Paragraph>}
      {lines.deductions[2].line && <Paragraph id="deposit-cancellation-deduction-3">{lines.deductions[2].line}</Paragraph>}
      {lines.deductions[3].line && <Paragraph id="deposit-cancellation-deduction-4">{lines.deductions[3].line}</Paragraph>}
      {lines.deductions[4].line && <Paragraph id="deposit-cancellation-deduction-5">{lines.deductions[4].line}</Paragraph>}
      {lines.deductions[5].line && <Paragraph id="deposit-cancellation-deduction-6">{lines.deductions[5].line}</Paragraph>}
      {lines.refundable && <Paragraph id="deposit-cancellation-refundable">Refundable advance: {lines.refundable}</Paragraph>}
      {lines.retention && (
        <Paragraph id="deposit-cancellation-retention">{lines.retentionLabel}: {lines.retention}</Paragraph>
      )}
      {lines.nonRefundable && (
        <Paragraph id="deposit-cancellation-non-refundable">Non-refundable advance, retained: {lines.nonRefundable}</Paragraph>
      )}
      <Paragraph id="deposit-cancellation-refunded">Refunded to you: {lines.refunded}</Paragraph>
      {lines.retained && <Paragraph id="deposit-cancellation-retained">Retained: {lines.retained}</Paragraph>}
      </>
    )}
  </>
);
