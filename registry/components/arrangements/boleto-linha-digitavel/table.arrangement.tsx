import { Cell, Paragraph, Row, Table, type Yield } from "docxcelerate/template";
import type { BoletoLinhaDigitavelLines } from "./lines.ts";

/**
 * Due date and amount in a label/value table. The typeable line stays out of
 * it: its 54 characters do not fit beside a label column at 11pt without
 * wrapping, and a wrapped line is the one a payer mistypes.
 */
export const table = (lines: BoletoLinhaDigitavelLines): Yield => (
  <>
    {lines.withheld && <Paragraph id="boleto-withheld">{lines.withheld}</Paragraph>}
    {/* 40mm holds "Valor do documento" at 11pt. */}
    {lines.line && (
      <Table id="boleto-details" columns={[{ width: 40 }, { width: "auto" }]}>
        {lines.dueDate && (
          <Row id="boleto-due-date">
            <Cell>Vencimento</Cell>
            <Cell>{lines.dueDate}</Cell>
          </Row>
        )}
        {lines.amount && (
          <Row id="boleto-amount">
            <Cell>Valor do documento</Cell>
            <Cell>{lines.amount}</Cell>
          </Row>
        )}
      </Table>
    )}
    {lines.line && <Paragraph id="boleto-line-label">Linha digitável</Paragraph>}
    {lines.line && <Paragraph id="boleto-line">{lines.line}</Paragraph>}
  </>
);
