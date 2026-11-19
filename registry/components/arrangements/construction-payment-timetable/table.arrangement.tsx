import { Cell, Paragraph, Row, Table, type Yield } from "docxcelerate/template";
import type { ConstructionPaymentTimetableLines } from "./lines.ts";

/**
 * The four dates in a label and date table, for a timetable read at a glance.
 * The default.
 */
export const table = (lines: ConstructionPaymentTimetableLines): Yield => (
  <>
    {lines.notApplicable && <Paragraph id="payment-timetable-not-applicable">{lines.notApplicable}</Paragraph>}
    {lines.applies && (
      <>
        <Paragraph id="payment-timetable-basis">{lines.basis}</Paragraph>
        {/* 50mm holds the longest label, "Pay less notice due by", on one line. */}
        <Table id="payment-timetable-dates" columns={[{ width: 50 }, { width: "auto" }]}>
          <Row id="payment-timetable-due">
            <Cell>Payment due date</Cell>
            <Cell>{lines.dueDate}</Cell>
          </Row>
          <Row id="payment-timetable-notice-by">
            <Cell>Payment notice due by</Cell>
            <Cell>{lines.paymentNoticeBy}</Cell>
          </Row>
          <Row id="payment-timetable-final">
            <Cell>Final date for payment</Cell>
            <Cell>{lines.finalDate}</Cell>
          </Row>
          <Row id="payment-timetable-pay-less-by">
            <Cell>Pay less notice due by</Cell>
            <Cell>{lines.payLessBy}</Cell>
          </Row>
        </Table>
        {lines.noticeLate && <Paragraph id="payment-timetable-notice-late">{lines.noticeLate}</Paragraph>}
        {lines.noticeGiven && <Paragraph id="payment-timetable-notice-given">{lines.noticeGiven}</Paragraph>}
        {lines.applicationStands && (
          <Paragraph id="payment-timetable-application-stands">{lines.applicationStands}</Paragraph>
        )}
        {lines.defaultNotice && <Paragraph id="payment-timetable-default-notice">{lines.defaultNotice}</Paragraph>}
        {lines.defaultOpen && <Paragraph id="payment-timetable-default-open">{lines.defaultOpen}</Paragraph>}
        <Paragraph id="payment-timetable-reckoning">{lines.reckoning}</Paragraph>
      </>
    )}
  </>
);
