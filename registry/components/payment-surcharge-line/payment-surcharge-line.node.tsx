import { arrange, type Nodes, useDeriver, useState } from "docxcelerate/template";
import type { PaymentSurchargeLineLines } from "../arrangements/payment-surcharge-line/lines.ts";
import { stacked } from "../arrangements/payment-surcharge-line/stacked.arrangement.tsx";
import { table } from "../arrangements/payment-surcharge-line/table.arrangement.tsx";
import paymentSurchargeLine from "../derivers/payment-surcharge-line.deriver.ts";

/**
 * The surcharge for the customer's chosen payment method, printed only where
 * charging it is lawful, never above a cap, taxed the way the main supply is
 * taxed, and with the disclosure the rules require. Otherwise no surcharge
 * and a warning for whoever issues the document. No heading, so it can sit
 * straight under the invoice totals.
 *
 * Decisions already made for you:
 *
 * - Banned cases print nothing but the warning: no UK consumer for any method;
 *   no consumer card in the UK or EU; no SEPA transfer in the EU; in the US
 *   only credit cards (commercial included), and never in a state that bans it.
 * - Every regime caps the surcharge at the cost of acceptance; with no cost
 *   given, nothing is charged. The US adds the network's and the state's cap.
 * - Every percentage, the request and each cap, is applied to the net total.
 *   That is the smallest base any of these rules could mean, so a cap stated
 *   as a share of "the transaction" holds whichever base a regulator reads.
 * - Rounded once, in minor units, and down at a cap so the cap is never
 *   exceeded.
 * - In the UK and EU the surcharge is further consideration for the main
 *   supply (Bookit, C-607/14): not exempt, and not a flat 20% line. It is
 *   spread across the invoice's VAT bands in proportion to their net, by
 *   largest remainder, and taxed at each band's rate, so a zero-rated band's
 *   share carries no VAT. In the US it is part of the sales price unless
 *   `usTaxable` says the state excludes it.
 * - A reduced rate after capping gets no warning: the label and the
 *   disclosure print the rate actually charged.
 *
 * Every rule is arithmetic or comparison on request data, which cannot be
 * written inline in a published document, so all of it is one deriver. This
 * node and its arrangements only place what the deriver returns.
 *
 * Not modelled: member-state bans beyond the EU defaults, other US states,
 * card-network rules other than the percentage cap, fixed-fee surcharges,
 * credit notes. The warning prints in the document, to recipients as well as
 * the issuer, because the framework has no author-only output: send `percent`
 * only for customers you mean to surcharge.
 *
 * Installed by `dxcl add payment-surcharge-line`, which also installs its
 * deriver and arrangements.
 */

/** What this component reads. Add these fields to your document data type. */
export interface PaymentSurchargeLineData {
  surcharge: {
    /**
     * Whose surcharging rules apply. `EU` means the Directive and Regulation
     * defaults; a member state that bans more is not modelled. Required:
     * anything else throws.
     */
    jurisdiction: "GB" | "EU" | "US";
    /**
     * The US state the sale is made in, as a two-letter postal code: `"NY"`.
     * Trimmed and upper-cased before it is compared. Read only for `US`.
     * Missing or empty there prints the "not applied" warning, because the
     * state decides whether a surcharge is allowed at all.
     */
    usState?: string;
    /**
     * Who is paying. Read only for `GB`, where any consumer surcharge is
     * banned; in the EU and US the method decides. Missing means `consumer`,
     * the reading that bans more. Any other value throws.
     */
    customer?: "consumer" | "business";
    /** How the customer is paying. Required. */
    method: {
      /**
       * The payment method. `credit-card`, `debit-card` and `prepaid-card`
       * are consumer cards; `commercial-card` is a business card. In `GB`,
       * `sepa` is treated as `bank-transfer`. Required: anything else throws.
       */
      kind: "credit-card" | "debit-card" | "prepaid-card" | "commercial-card" | "sepa" | "bank-transfer";
      /**
       * The card network, lower case: `"visa"`, `"mastercard"`. Read only in
       * the US, for the network cap. Missing or unknown gets the lowest cap
       * in the table (3%).
       */
      network?: string;
    };
    /**
     * The surcharge rate the seller asks for, as a percentage of the net
     * total: `2.5` means 2.5%. Rounded to 3 decimal places. Missing or 0
     * prints nothing at all and no warning. Negative or not a finite number
     * throws.
     */
    percent?: number;
    /**
     * What accepting this method costs the seller, as a percentage of the
     * transaction: the merchant service charge rate, such as `1.95`. Every
     * regime caps the surcharge here. Missing, when a surcharge would
     * otherwise be charged, prints the "not applied" warning. Negative or not
     * finite throws.
     */
    costOfAcceptance?: number;
    /**
     * US only: whether the state taxes the surcharge as part of the sales
     * price. Missing means true, because most states' definition of sales
     * price includes charges needed to complete the sale. False prints no
     * tax on the surcharge.
     */
    usTaxable?: boolean;
    /**
     * ISO 4217 code: `"GBP"`. Sets the minor units (JPY has none) and how
     * amounts print. Required: missing or unknown throws.
     */
    currency: string;
    /**
     * The invoice's totals by tax rate, before any surcharge, as the totals
     * section already shows them. One entry per rate. The surcharge is
     * spread across these by net. An empty list gives a zero base and prints
     * nothing.
     */
    bands: Array<{
      /**
       * The VAT or sales-tax rate, as a percentage: `20`, `5`, `0`, `8.875`.
       * An exempt or zero-rated band is `0`. Negative or not finite throws.
       */
      ratePercent: number;
      /**
       * The band's net amount, in major units. Negative or not finite throws:
       * a credit note is never surcharged.
       */
      net: number;
      /**
       * The tax already charged on that net, in major units. Read only to
       * work out the pre-surcharge gross for the total line, and never
       * recomputed. Missing counts as 0. Negative or not finite throws.
       */
      vat?: number;
    }>;
  };
}

/**
 * Where the lines sit.
 *
 * - `table`: the surcharge, its tax and the new total in a label/amount table
 *   under the invoice totals, then the notes. The default.
 * - `stacked`: each line as a sentence, one under another, down the text column.
 */
export type PaymentSurchargeLineArrangement = "table" | "stacked";

/** Props for {@linkcode PaymentSurchargeLine}. */
export interface PaymentSurchargeLineProps {
  /** Where the lines sit. Defaults to `"table"`. */
  arrangement?: PaymentSurchargeLineArrangement;
}

export const PaymentSurchargeLine: Nodes<PaymentSurchargeLineProps> = async ({ arrangement }) => {
  const [surcharge] = useState((data: PaymentSurchargeLineData) => data.surcharge);
  const out = await useDeriver(paymentSurchargeLine, [surcharge]);

  const lines: PaymentSurchargeLineLines = {
    surchargeLabel: out.surchargeLabel,
    surchargeAmount: out.surchargeAmount,
    taxLabel: out.taxLabel,
    taxAmount: out.taxAmount,
    taxNote: out.taxNote,
    totalLabel: out.totalLabel,
    totalAmount: out.totalAmount,
    disclosure: out.disclosure,
    warning: out.warning,
  };

  return <>{arrange(arrangement, { table: () => table(lines), stacked: () => stacked(lines) })}</>;
};
