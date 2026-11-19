import { Paragraph, type Yield } from "docxcelerate/template";
import type { BoletoLinhaDigitavelLines } from "./lines.ts";

/**
 * Paragraphs down the column, the typeable line on its own beneath the
 * labelled due date and amount. The default: it reads in any width.
 */
export const stacked = (lines: BoletoLinhaDigitavelLines): Yield => (
  <>
    {lines.withheld && <Paragraph id="boleto-withheld">{lines.withheld}</Paragraph>}
    {lines.dueDate && <Paragraph id="boleto-due-date">Vencimento: {lines.dueDate}</Paragraph>}
    {lines.amount && <Paragraph id="boleto-amount">Valor do documento: {lines.amount}</Paragraph>}
    {lines.line && <Paragraph id="boleto-line-label">Linha digitável</Paragraph>}
    {lines.line && <Paragraph id="boleto-line">{lines.line}</Paragraph>}
  </>
);
