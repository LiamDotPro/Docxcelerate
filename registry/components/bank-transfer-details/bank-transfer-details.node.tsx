import { arrange, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import type { BankTransferDetailsLines } from "../arrangements/bank-transfer-details/lines.ts";
import { split } from "../arrangements/bank-transfer-details/split.arrangement.tsx";
import { stacked } from "../arrangements/bank-transfer-details/stacked.arrangement.tsx";
import { table } from "../arrangements/bank-transfer-details/table.arrangement.tsx";
import bankTransferDetails from "../derivers/bank-transfer-details.deriver.ts";

/**
 * One way to pay this invoice by bank transfer: the one that suits this payer,
 * with only details that have passed their checks.
 *
 * One rail, not all of them: a payer shown a sort code, an IBAN and a routing
 * number picks the wrong one. The payer's country and the invoice currency
 * decide it: sort code and account for a UK payer in sterling, IBAN and BIC for
 * SEPA in euros, ABA routing and account for a US payer in dollars, and IBAN and
 * BIC with a SWIFT note for anyone else.
 *
 * Checked, then printed: a transposed digit would only be caught at the payer's
 * bank, or not at all. A detail that fails is withheld and the payer falls to
 * the next rail that can be printed honestly, and to a "call us" line when
 * there is none. A rail is never printed half-filled.
 *
 * The reference is fitted to the rail: banks cut a longer one short without
 * saying so and the payment cannot then be matched. When the invoice reference
 * does not fit, the customer reference is printed instead. Never a truncated
 * one: when neither fits, the document fails.
 *
 * The account name is printed as given, because UK Confirmation of Payee
 * compares it character for character.
 *
 * When the details changed recently, a fixed warning to confirm them by phone
 * comes first, with no number in it: a forged invoice would print the
 * fraudster's. It is never generated text.
 *
 * All of this is decided once, in one deriver; arithmetic and string work on
 * request data throw at publish time without one. This file and the
 * arrangements, in `arrangements/bank-transfer-details/`, only place the lines.
 *
 * Installed by `dxcl add bank-transfer-details`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface BankTransferDetailsData {
  /** Who is being paid, and every account they can be paid into. */
  payee: {
    /**
     * The account name exactly as the bank holds it. Only outer whitespace is
     * trimmed: no case change, no abbreviation, no "Ltd" to "Limited". UK
     * Confirmation of Payee compares it character for character. Required: the
     * deriver throws when it is missing or blank.
     */
    accountName: string;
    /**
     * UK domestic details, used for a UK payer paying in GBP. Spaces and
     * hyphens are ignored. Missing, or either part invalid, means the UK rail
     * is not offered and the payer is sent to IBAN and BIC instead.
     */
    uk?: {
      /** Six digits, in any of `601613`, `60-16-13` or `60 16 13`. Printed `60-16-13`. */
      sortCode: string;
      /** Exactly eight digits. A shorter UK account number must be zero-padded by the caller. */
      accountNumber: string;
    };
    /**
     * The IBAN, with or without spaces, in any case. Checked with ISO 13616
     * mod-97 and printed upper case in groups of four. Missing or invalid
     * means no SEPA and no international rail.
     */
    iban?: string;
    /**
     * The BIC (SWIFT code), 8 or 11 characters. Printed upper case, without
     * spaces, as given (an 8-character BIC is not padded with XXX). Optional
     * for SEPA: the IBAN alone is printed when it is missing or invalid.
     * Required for the international rail: without a valid one, that rail is
     * not offered.
     */
    bic?: string;
    /**
     * US details, used for a US payer paying in USD. Spaces and hyphens are
     * ignored. Missing, or an invalid ACH routing or account number, means the
     * US rail is not offered.
     */
    us?: {
      /** The ABA routing number for ACH: nine digits that pass the 3-7-1 checksum. */
      achRouting: string;
      /**
       * The routing number for wires, when the bank uses a different one.
       * Missing, or the same as `achRouting`, means one routing number serves
       * both. Invalid means it is withheld and only ACH is offered.
       */
      wireRouting?: string;
      /** The account number: 4 to 17 digits. */
      accountNumber: string;
    };
    /**
     * The ISO 8601 date these details changed, when they differ from the
     * details on this payer's previous invoice. Set it for invoices issued
     * while the change is new, and remove it once the payer has paid to the
     * new account: the component cannot see invoice history. When it is
     * present, the verify-by-phone warning is printed above the details. A
     * value that is not YYYY-MM-DD still prints the warning, without the date.
     * Missing or empty means no warning.
     */
    detailsChangedOn?: string;
  };
  /** Who is paying. */
  payer: {
    /**
     * ISO 3166-1 alpha-2 country code, any case. `UK` is accepted as `GB`.
     * Missing or unrecognised means the payer is treated as international.
     */
    country?: string;
  };
  /** The invoice being paid. */
  invoice: {
    /**
     * ISO 4217 code, any case: `GBP`, `EUR`, `USD`. Together with the country,
     * it picks the rail. Required: the deriver throws unless it is three
     * letters.
     */
    currency: string;
    /**
     * The reference the payer should quote, usually the invoice number.
     * Printed when it fits the chosen rail. Required: the deriver throws when
     * it is missing or blank.
     */
    reference: string;
    /**
     * A shorter reference for when `reference` does not fit the rail, such as
     * the customer's account number. Printed only in that case. When neither
     * reference fits, the deriver throws rather than print one the bank will
     * cut short.
     */
    customerReference?: string;
  };
}

/**
 * Where the lines sit.
 *
 * - `stacked`: one labelled line under another. The default.
 * - `table`: labels in a narrow column, values beside them.
 * - `split`: where to send the money on the left, the reference to quote ranged right.
 */
export type BankTransferDetailsArrangement = "stacked" | "table" | "split";

/** Props for {@linkcode BankTransferDetails}. */
export interface BankTransferDetailsProps {
  /** Where the lines sit. Defaults to `"stacked"`. */
  arrangement?: BankTransferDetailsArrangement;
}

export const BankTransferDetails: SectionComponent<BankTransferDetailsProps> = async ({ arrangement }) => {
  const [payee] = useState((data: BankTransferDetailsData) => data.payee);
  const [payer] = useState((data: BankTransferDetailsData) => data.payer);
  const [invoice] = useState((data: BankTransferDetailsData) => data.invoice);
  const details = await useDeriver(bankTransferDetails, [payee, payer, invoice]);

  const lines: BankTransferDetailsLines = {
    warning: details.warning,
    method: details.method,
    accountName: details.accountName,
    sortCode: details.sortCode,
    routingLabel: details.routingLabel,
    routing: details.routing,
    wireRouting: details.wireRouting,
    accountNumber: details.accountNumber,
    iban: details.iban,
    bic: details.bic,
    reference: details.reference,
  };

  return (
    <Section id="bank-transfer" title="How to pay">
      {arrange(arrangement, {
        stacked: () => stacked(lines),
        table: () => table(lines),
        split: () => split(lines),
      })}
    </Section>
  );
};
