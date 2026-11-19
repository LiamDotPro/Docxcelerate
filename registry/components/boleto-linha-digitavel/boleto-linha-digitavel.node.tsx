import { arrange, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import type { BoletoLinhaDigitavelLines } from "../arrangements/boleto-linha-digitavel/lines.ts";
import { stacked } from "../arrangements/boleto-linha-digitavel/stacked.arrangement.tsx";
import { table } from "../arrangements/boleto-linha-digitavel/table.arrangement.tsx";
import boletoLinhaDigitavel from "../derivers/boleto-linha-digitavel.deriver.ts";

/**
 * The typeable line of a Brazilian boleto, built from the invoice or checked
 * against the one the bank returned.
 *
 * The due-date factor rolled over on 2025-02-22. It counts days from
 * 1997-10-07 up to 9999 on 2025-02-21, then restarts at 1000, so it is never
 * "days since 1997". The general check digit is modulus 11 over the barcode
 * order. The three field check digits are modulus 10 over the line order.
 * Those two orders are not the same, and mixing them up is the classic bug.
 *
 * If the bank's line disagrees with the invoice in any digit, the document
 * fails. The bank collects what its line says, so printing a line for a
 * different amount or date is worse than printing none. If a line cannot be
 * built at all, it is withheld with a visible reason, never approximated. This
 * happens for an invoice not in reais, a bank code or free field of the wrong
 * length, a due date before the factor window, or nothing to pay.
 *
 * The drawn ITF barcode is not produced. A payer can pay by the typeable line
 * alone. Labels are Portuguese because a boleto is paid through a Brazilian
 * bank; thrown errors are English because they are for the developer.
 *
 * All of it is one deriver: dates and amounts from request data cannot be
 * computed at publish time.
 *
 * Arrangement: where the lines sit is chosen with the `arrangement` prop. This
 * file decides every line. Each arrangement in
 * `arrangements/boleto-linha-digitavel/` only places the lines.
 *
 * Installed by `dxcl add boleto-linha-digitavel`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface BoletoData {
  boleto: {
    /**
     * The issuing bank's three-digit FEBRABAN code: "001" Banco do Brasil,
     * "237" Bradesco, "341" Itaú. Trimmed, then it must be exactly three
     * digits. Missing or anything else withholds the line.
     */
    bankCode: string;
    /**
     * The bank-specific 25-digit free field (campo livre), as your bank's
     * layout assembles it from the nosso número, agency, account and wallet.
     * Trimmed, then it must be exactly 25 digits, with no dots, spaces or
     * hyphens inside. Missing or any other length withholds the line. This
     * component does not assemble it: every bank lays it out differently.
     */
    freeField: string;
    /**
     * The linha digitável the bank returned when the boleto was registered.
     * Spaces and dots are ignored, then it must be 47 digits, or the
     * document fails. When given, it is checked against the line built from
     * the invoice. Any difference fails the document, naming what differs.
     * Missing or empty means the line is built from the invoice and not
     * cross-checked. It is ignored (not checked) when the line is withheld.
     */
    line?: string;
  };
  invoice: {
    /**
     * The amount to collect, in reais. Rounded half up to whole centavos as
     * written (1.005 becomes R$ 1,01). Not a finite number, below one
     * centavo after rounding, or over R$ 99.999.999,99 withholds the line.
     */
    total: number;
    /**
     * The due date, as an ISO 8601 date: `2026-12-15`. Anything after the
     * tenth character is ignored. Missing or empty means no due date: factor
     * 0000, printed "Contra apresentação". A date before 2000-07-03 withholds
     * the line. A value that is not a real YYYY-MM-DD date (`2026-02-30`,
     * `15/12/2026`) fails the document.
     */
    dueDate?: string;
    /**
     * Must be "BRL": a boleto is payable in reais only. Missing or anything
     * else withholds the line; it is never defaulted.
     */
    currency: "BRL";
  };
}

/**
 * Where the lines sit.
 *
 * - `stacked`: paragraphs down the column. The default.
 * - `table`: due date and amount in a label/value table, the typeable line
 *   full width beneath it.
 */
export type BoletoLinhaDigitavelArrangement = "stacked" | "table";

/** Props for {@linkcode BoletoLinhaDigitavel}. */
export interface BoletoLinhaDigitavelProps {
  /** Where the lines sit. Defaults to `"stacked"`. */
  arrangement?: BoletoLinhaDigitavelArrangement;
}

export const BoletoLinhaDigitavel: SectionComponent<BoletoLinhaDigitavelProps> = async ({ arrangement }) => {
  const [boleto] = useState((data: BoletoData) => data.boleto);
  const [invoice] = useState((data: BoletoData) => data.invoice);
  const result = await useDeriver(boletoLinhaDigitavel, [boleto, invoice]);

  const lines: BoletoLinhaDigitavelLines = {
    line: result.line,
    dueDate: result.dueDate,
    amount: result.amount,
    withheld: result.withheld,
  };

  return (
    <Section id="boleto" title="Boleto bancário">
      {arrange(arrangement, {
        stacked: () => stacked(lines),
        table: () => table(lines),
      })}
    </Section>
  );
};
