import { Cell, Paragraph, Row, Table, type Yield } from "docxcelerate/template";
import type { BankTransferDetailsLines } from "./lines.ts";

/**
 * Labels in a narrow column with the values beside them. A payer copying digits
 * into a banking app finds every value starting at the same place, and an IBAN
 * or routing number is never broken by a label running into it.
 */
export const table = (lines: BankTransferDetailsLines): Yield => (
  <>
    {lines.warning && <Paragraph id="bank-transfer-warning">{lines.warning}</Paragraph>}
    <Paragraph id="bank-transfer-method">{lines.method}</Paragraph>
    {/* 45mm holds the longest label, "Routing number (ACH and wire)", on one line. */}
    {/* On the none rail no row prints, and a row-less table is not valid DOCX. */}
    {lines.accountName && <Table id="bank-transfer-details" columns={[{ width: 45 }, { width: "auto" }]}>
      {lines.accountName && (
        <Row id="bank-transfer-account-name">
          <Cell>Account name</Cell>
          <Cell>{lines.accountName}</Cell>
        </Row>
      )}
      {lines.sortCode && (
        <Row id="bank-transfer-sort-code">
          <Cell>Sort code</Cell>
          <Cell>{lines.sortCode}</Cell>
        </Row>
      )}
      {lines.routing && (
        <Row id="bank-transfer-routing">
          <Cell>{lines.routingLabel}</Cell>
          <Cell>{lines.routing}</Cell>
        </Row>
      )}
      {lines.wireRouting && (
        <Row id="bank-transfer-wire-routing">
          <Cell>Wire routing number</Cell>
          <Cell>{lines.wireRouting}</Cell>
        </Row>
      )}
      {lines.accountNumber && (
        <Row id="bank-transfer-account-number">
          <Cell>Account number</Cell>
          <Cell>{lines.accountNumber}</Cell>
        </Row>
      )}
      {lines.iban && (
        <Row id="bank-transfer-iban">
          <Cell>IBAN</Cell>
          <Cell>{lines.iban}</Cell>
        </Row>
      )}
      {lines.bic && (
        <Row id="bank-transfer-bic">
          <Cell>BIC</Cell>
          <Cell>{lines.bic}</Cell>
        </Row>
      )}
      {lines.reference && (
        <Row id="bank-transfer-reference">
          <Cell>Payment reference</Cell>
          <Cell>{lines.reference}</Cell>
        </Row>
      )}
    </Table>}
  </>
);
