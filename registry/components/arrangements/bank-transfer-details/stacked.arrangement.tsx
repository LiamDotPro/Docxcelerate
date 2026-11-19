import { Paragraph, type Yield } from "docxcelerate/template";
import type { BankTransferDetailsLines } from "./lines.ts";

/**
 * One labelled line under another, down the text column. The default: it
 * reads at any width, and nothing has to line up with anything.
 */
export const stacked = (lines: BankTransferDetailsLines): Yield => (
  <>
    {lines.warning && <Paragraph id="bank-transfer-warning">{lines.warning}</Paragraph>}
    <Paragraph id="bank-transfer-method">{lines.method}</Paragraph>
    {lines.accountName && <Paragraph id="bank-transfer-account-name">Account name: {lines.accountName}</Paragraph>}
    {lines.sortCode && <Paragraph id="bank-transfer-sort-code">Sort code: {lines.sortCode}</Paragraph>}
    {lines.routing && <Paragraph id="bank-transfer-routing">{lines.routingLabel}: {lines.routing}</Paragraph>}
    {lines.wireRouting && (
      <Paragraph id="bank-transfer-wire-routing">Wire routing number: {lines.wireRouting}</Paragraph>
    )}
    {lines.accountNumber && (
      <Paragraph id="bank-transfer-account-number">Account number: {lines.accountNumber}</Paragraph>
    )}
    {lines.iban && <Paragraph id="bank-transfer-iban">IBAN: {lines.iban}</Paragraph>}
    {lines.bic && <Paragraph id="bank-transfer-bic">BIC: {lines.bic}</Paragraph>}
    {lines.reference && <Paragraph id="bank-transfer-reference">Payment reference: {lines.reference}</Paragraph>}
  </>
);
