import { Paragraph, type Yield } from "docxcelerate/template";
import type { ConstructionPaymentTimetableLines } from "./lines.ts";

/** One line per date, for a letter that reads as prose. */
export const stacked = (lines: ConstructionPaymentTimetableLines): Yield => (
  <>
    {lines.notApplicable && <Paragraph id="payment-timetable-not-applicable">{lines.notApplicable}</Paragraph>}
    {lines.applies && (
      <>
        <Paragraph id="payment-timetable-basis">{lines.basis}</Paragraph>
        <Paragraph id="payment-timetable-due">Payment due date: {lines.dueDate}</Paragraph>
        <Paragraph id="payment-timetable-notice-by">Payment notice due by: {lines.paymentNoticeBy}</Paragraph>
        <Paragraph id="payment-timetable-final">Final date for payment: {lines.finalDate}</Paragraph>
        <Paragraph id="payment-timetable-pay-less-by">Pay less notice due by: {lines.payLessBy}</Paragraph>
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
