import { Cell, Paragraph, Row, Table, type Yield } from "docxcelerate/template";
import type { BankTransferDetailsLines } from "./lines.ts";

/**
 * Where to send the money on the left, the reference to quote ranged right. The
 * payer enters the reference in a separate field from the account details, and
 * on many invoices this block faces the totals.
 */
export const split = (lines: BankTransferDetailsLines): Yield => (
  <>
    {lines.warning && <Paragraph id="bank-transfer-warning">{lines.warning}</Paragraph>}
    <Paragraph id="bank-transfer-method">{lines.method}</Paragraph>
    <Table id="bank-transfer-details" columns={[{ width: "auto" }, { width: 60, align: "right" }]}>
      <Row id="bank-transfer-details-row">
        <Cell>
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
        </Cell>
        <Cell>
          {lines.reference && (
            <Paragraph id="bank-transfer-reference">Quote this reference: {lines.reference}</Paragraph>
          )}
        </Cell>
      </Row>
    </Table>
  </>
);
