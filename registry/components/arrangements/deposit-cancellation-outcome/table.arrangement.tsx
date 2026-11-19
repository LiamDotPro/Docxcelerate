import { Cell, Row, Table, type Yield } from "docxcelerate/template";
import type { DepositCancellationOutcomeLines } from "./lines.ts";

/**
 * Each label beside its amount, amounts right-aligned. The default: the
 * section is arithmetic the reader checks down a column, paid, less what is
 * retained, equals what is refunded.
 */
export const table = (lines: DepositCancellationOutcomeLines): Yield => (
  <>
    {lines.paid && (
      /* Amounts get what a figure like £1,284,750.00 needs; descriptions, which run long, get the rest. */
      <Table id="deposit-cancellation-amounts" columns={[{ width: "auto" }, { width: 40, align: "right" }]}>
      <Row id="deposit-cancellation-paid">
        <Cell>Paid in advance</Cell>
        <Cell>{lines.paid}</Cell>
      </Row>
      {lines.security && (
        <Row id="deposit-cancellation-security">
          <Cell>Security deposit</Cell>
          <Cell>{lines.security}</Cell>
        </Row>
      )}
      {lines.deductions[0].amount && (
        <Row id="deposit-cancellation-deduction-1">
          <Cell>{lines.deductions[0].description}</Cell>
          <Cell>{lines.deductions[0].amount}</Cell>
        </Row>
      )}
      {lines.deductions[1].amount && (
        <Row id="deposit-cancellation-deduction-2">
          <Cell>{lines.deductions[1].description}</Cell>
          <Cell>{lines.deductions[1].amount}</Cell>
        </Row>
      )}
      {lines.deductions[2].amount && (
        <Row id="deposit-cancellation-deduction-3">
          <Cell>{lines.deductions[2].description}</Cell>
          <Cell>{lines.deductions[2].amount}</Cell>
        </Row>
      )}
      {lines.deductions[3].amount && (
        <Row id="deposit-cancellation-deduction-4">
          <Cell>{lines.deductions[3].description}</Cell>
          <Cell>{lines.deductions[3].amount}</Cell>
        </Row>
      )}
      {lines.deductions[4].amount && (
        <Row id="deposit-cancellation-deduction-5">
          <Cell>{lines.deductions[4].description}</Cell>
          <Cell>{lines.deductions[4].amount}</Cell>
        </Row>
      )}
      {lines.deductions[5].amount && (
        <Row id="deposit-cancellation-deduction-6">
          <Cell>{lines.deductions[5].description}</Cell>
          <Cell>{lines.deductions[5].amount}</Cell>
        </Row>
      )}
      {lines.refundable && (
        <Row id="deposit-cancellation-refundable">
          <Cell>Refundable advance</Cell>
          <Cell>{lines.refundable}</Cell>
        </Row>
      )}
      {lines.retention && (
        <Row id="deposit-cancellation-retention">
          <Cell>{lines.retentionLabel}</Cell>
          <Cell>{lines.retention}</Cell>
        </Row>
      )}
      {lines.nonRefundable && (
        <Row id="deposit-cancellation-non-refundable">
          <Cell>Non-refundable advance, retained</Cell>
          <Cell>{lines.nonRefundable}</Cell>
        </Row>
      )}
      <Row id="deposit-cancellation-refunded">
        <Cell>Refunded to you</Cell>
        <Cell>{lines.refunded}</Cell>
      </Row>
      {lines.retained && (
        <Row id="deposit-cancellation-retained">
          <Cell>Retained</Cell>
          <Cell>{lines.retained}</Cell>
        </Row>
      )}
      </Table>
    )}
  </>
);
