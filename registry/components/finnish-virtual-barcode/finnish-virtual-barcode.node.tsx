import { arrange, type Nodes, useDeriver, useState } from "docxcelerate/template";
import type { FinnishVirtualBarcodeLines } from "../arrangements/finnish-virtual-barcode/lines.ts";
import { stacked } from "../arrangements/finnish-virtual-barcode/stacked.arrangement.tsx";
import { table } from "../arrangements/finnish-virtual-barcode/table.arrangement.tsx";
import finnishVirtualBarcode from "../derivers/finnish-virtual-barcode.deriver.ts";

/**
 * The payee's IBAN, the payment reference and the 54-digit virtual bank
 * barcode (virtuaaliviivakoodi) a Finnish payer pastes into a banking app. A
 * truncated or mis-versioned code would scan into the wrong account and a
 * payer cannot spot that by eye, so the code is printed only when every check
 * passes.
 *
 * The reference decides the version: an RF reference gives version 5, a
 * viitenumero version 4. There is no flag, because a flag could disagree with
 * the reference. Version 2 is never produced.
 *
 * Any doubt withholds the code: a bad IBAN, a non-Finnish one, a currency
 * other than EUR, a bad reference check, an alphanumeric RF, a bad due date,
 * or an ERP code that differs from the rebuilt one. The IBAN and reference
 * are still printed, so the invoice can be paid by hand. A total of zero or
 * less prints nothing at all.
 *
 * All of it is one deriver, since it is checksums and string surgery on
 * request data, which publishing refuses. The node computes nothing; the
 * arrangements guard each line on its string. There is no heading, so it can
 * sit under an invoice's totals or giro block.
 *
 * Installed by `dxcl add finnish-virtual-barcode`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface FinnishVirtualBarcodeData {
  payee: {
    /**
     * The account to be paid, as an IBAN. Spaces are ignored and letters are
     * upper-cased, so "fi79 4405 2020 0360 82" is fine. It must pass the
     * IBAN mod-97 check and start with FI: the guide allows a barcode only
     * for a Finnish account. A non-Finnish IBAN that is otherwise valid is
     * still printed, without a code. An invalid one is printed as given
     * (trimmed), without a code; the payer's bank refuses it at payment
     * rather than sending money anywhere. Missing or empty withholds the
     * code and prints no IBAN line.
     */
    iban: string;
  };
  invoice: {
    /**
     * ISO 4217 code. Must be "EUR" (compared after trimming and
     * upper-casing). Anything else, or missing, withholds the code: the
     * barcode's amount field is euros by definition.
     */
    currency: string;
    /**
     * What the payer owes, in euros, to the cent: 4883.15. Zero or less
     * (a credit note, an invoice already paid) prints nothing at all. Over
     * 999 999,99 the amount field is zero-filled so the payer types it in.
     * Not a finite number, or finer than a cent (12.345), withholds the code,
     * because a barcode amount a cent off the printed total is a wrong
     * barcode. Not read when `openAmount` is true.
     */
    total: number;
    /**
     * The due date, ISO 8601 (YYYY-MM-DD; anything after the tenth character
     * is ignored). Written into the code as YYMMDD. Missing or empty gives
     * 000000. A value that is not a real date in 2000-2099 withholds the
     * code. Ignored when `onReceipt` is true.
     */
    dueDate?: string;
    /**
     * True when the invoice is payable on receipt. The code's due-date field
     * becomes 000000 rather than a guessed date. Missing means false.
     */
    onReceipt?: boolean;
    /**
     * True when the payer decides the amount (a donation, an advance). The
     * amount field is zero-filled and `total` is not read. Missing means false.
     */
    openAmount?: boolean;
  };
  payment: {
    /**
     * The payment reference: a Finnish viitenumero ("86851 62596 19897") or
     * an ISO 11649 RF reference ("RF09 8685 1625 9619 897"). Spaces are
     * ignored. A reference starting with RF (any case) is read as RF and
     * gives a version 5 code; anything else is read as a viitenumero and
     * gives version 4. A reference that fails its checks is printed as given
     * (trimmed), without a code. Missing or empty withholds the code and
     * prints no reference line.
     */
    reference: string;
  };
  /**
   * The 54-digit code the ERP already issued, if it issues one. Spaces are
   * ignored. It is never printed on trust: the code is rebuilt from the
   * fields above, and if the two differ the code is withheld, because
   * nobody can say which one is wrong, and the wrong one pays the wrong
   * account or the wrong amount. Missing or empty means "build it".
   */
  virtualBarcode?: string;
}

/**
 * Where the lines sit.
 *
 * - `stacked`: each labelled line under the last, the code on its own line
 *   under its label. The default.
 * - `table`: the IBAN and reference as label/value rows, the code under the
 *   table at full column width.
 */
export type FinnishVirtualBarcodeArrangement = "stacked" | "table";

/** Props for {@linkcode FinnishVirtualBarcode}. */
export interface FinnishVirtualBarcodeProps {
  /** Where the lines sit. Defaults to `"stacked"`. */
  arrangement?: FinnishVirtualBarcodeArrangement;
}

export const FinnishVirtualBarcode: Nodes<FinnishVirtualBarcodeProps> = async ({ arrangement }) => {
  const [payee] = useState((data: FinnishVirtualBarcodeData) => data.payee);
  const [invoice] = useState((data: FinnishVirtualBarcodeData) => data.invoice);
  const [payment] = useState((data: FinnishVirtualBarcodeData) => data.payment);
  const [virtualBarcode] = useState((data: FinnishVirtualBarcodeData) => data.virtualBarcode);
  const v = await useDeriver(finnishVirtualBarcode, [payee, invoice, payment, virtualBarcode]);

  const lines: FinnishVirtualBarcodeLines = { iban: v.iban, reference: v.reference, code: v.code };

  return <>{arrange(arrangement, { stacked: () => stacked(lines), table: () => table(lines) })}</>;
};
