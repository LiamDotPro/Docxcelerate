import { Paragraph, type Yield } from "docxcelerate/template";
import type { DunningReminderLines } from "./lines.ts";

/**
 * One sentence per invoice down the text column, then the total. Suits plain
 * letters and narrow columns, where four table columns would crowd.
 */
export const stacked = (lines: DunningReminderLines): Yield => (
  <>
    {lines.invoices[0].number && (
      <Paragraph id="dunning-reminder-invoice-1">{lines.invoices[0].line}</Paragraph>
    )}
    {lines.invoices[1].number && (
      <Paragraph id="dunning-reminder-invoice-2">{lines.invoices[1].line}</Paragraph>
    )}
    {lines.invoices[2].number && (
      <Paragraph id="dunning-reminder-invoice-3">{lines.invoices[2].line}</Paragraph>
    )}
    {lines.invoices[3].number && (
      <Paragraph id="dunning-reminder-invoice-4">{lines.invoices[3].line}</Paragraph>
    )}
    {lines.invoices[4].number && (
      <Paragraph id="dunning-reminder-invoice-5">{lines.invoices[4].line}</Paragraph>
    )}
    {lines.invoices[5].number && (
      <Paragraph id="dunning-reminder-invoice-6">{lines.invoices[5].line}</Paragraph>
    )}
    {lines.invoices[6].number && (
      <Paragraph id="dunning-reminder-invoice-7">{lines.invoices[6].line}</Paragraph>
    )}
    {lines.invoices[7].number && (
      <Paragraph id="dunning-reminder-invoice-8">{lines.invoices[7].line}</Paragraph>
    )}
    {lines.overflow && <Paragraph id="dunning-reminder-overflow">{lines.overflow}</Paragraph>}
    <Paragraph id="dunning-reminder-total">Total overdue: {lines.total}</Paragraph>
  </>
);
