import { Paragraph, Text, type Yield } from "docxcelerate/template";
import type { PatientResponsibilityBreakdownLines } from "./lines.ts";

/**
 * One labelled block per claim line, down the text column, then a total block.
 * For a narrow column, or a statement read on a phone as a PDF: nothing has to
 * line up with anything.
 */
export const stacked = (lines: PatientResponsibilityBreakdownLines): Yield => (
  <>
    {lines.rows.map((row) => (
      <>
        <Paragraph id="patient-responsibility-line-service"><Text bold>{row.service}</Text></Paragraph>
        <Paragraph id="patient-responsibility-line-detail">{row.detail}</Paragraph>
        <Paragraph id="patient-responsibility-line-billed">Billed: {row.billed}</Paragraph>
        <Paragraph id="patient-responsibility-line-allowed">Allowed amount: {row.allowed}</Paragraph>
        <Paragraph id="patient-responsibility-line-deductible">Deductible: {row.deductible}</Paragraph>
        <Paragraph id="patient-responsibility-line-copay">Copay: {row.copay}</Paragraph>
        <Paragraph id="patient-responsibility-line-coinsurance">Coinsurance: {row.coinsurance}</Paragraph>
        <Paragraph id="patient-responsibility-line-you-owe">You owe: {row.youOwe}</Paragraph>
        {row.note && <Paragraph id="patient-responsibility-line-note">{row.note}</Paragraph>}
      </>
    ))}
    <Paragraph id="patient-responsibility-total-heading"><Text bold>Total</Text></Paragraph>
    <Paragraph id="patient-responsibility-total-billed">Billed: {lines.total.billed}</Paragraph>
    <Paragraph id="patient-responsibility-total-allowed">Allowed amount: {lines.total.allowed}</Paragraph>
    <Paragraph id="patient-responsibility-total-deductible">Deductible: {lines.total.deductible}</Paragraph>
    <Paragraph id="patient-responsibility-total-copay">Copay: {lines.total.copay}</Paragraph>
    <Paragraph id="patient-responsibility-total-coinsurance">Coinsurance: {lines.total.coinsurance}</Paragraph>
    <Paragraph id="patient-responsibility-total-you-owe">You owe: {lines.total.youOwe}</Paragraph>
  </>
);
