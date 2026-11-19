import { arrange, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import type { InvoiceDocumentTypeLines } from "../arrangements/invoice-document-type/lines.ts";
import { split } from "../arrangements/invoice-document-type/split.arrangement.tsx";
import { stacked } from "../arrangements/invoice-document-type/stacked.arrangement.tsx";
import { table } from "../arrangements/invoice-document-type/table.arrangement.tsx";
import invoiceDocumentType from "../derivers/invoice-document-type.deriver.ts";

/**
 * The heading and opening lines of an invoice: what the document is allowed to
 * call itself, and which dates it has to carry. Those are decisions about the
 * law, not about layout.
 *
 * Not VAT-registered: a seller who is not registered gets "Invoice", with no
 * VAT number and no VAT wording. Calling the document a VAT invoice is an
 * offence in the UK. A missing registration flag is therefore read as "not
 * registered": the wrong answer that way costs the customer a reclaim, while
 * the other way round is unlawful.
 *
 * Registered: a registered seller gets "VAT invoice". It becomes "Simplified
 * VAT invoice" only when the seller allows simplified invoices and the
 * VAT-inclusive total is at or under the local ceiling.
 *
 * Pro forma: headed as one, and when VAT is in play it says "This is not a VAT
 * invoice." That stops the customer reclaiming on it.
 *
 * Self-billed: headed as one, and carries the statement its jurisdiction
 * requires.
 *
 * Date of supply: printed only when it falls on a different day from the issue
 * date, which is the UK and EU rule. In Germany a line is always printed,
 * because UStG section 14(4) no. 6 requires the Leistungsdatum even when it
 * equals the invoice date.
 *
 * All of this is decided once, in one deriver, and the node prints what the
 * deriver returns. Written as ordinary conditionals the cases come to about
 * sixteen published arms, half the default `branchLimit` of 32 spent before
 * the rest of the invoice, and dates from request data cannot be formatted at
 * publish time without a deriver anyway.
 *
 * Arrangement: where the lines sit is a separate choice from what they say,
 * made with the `arrangement` prop. This file decides every line; each
 * arrangement, in `arrangements/invoice-document-type/`, only places them, so
 * all of them print the same lines under the same conditions.
 *
 * Installed by `dxcl add invoice-document-type`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface InvoiceDocumentTypeData {
  invoice: {
    /**
     * The invoice number, unique in the seller's sequence. Printed as given.
     * Required: the deriver throws when it is missing or blank.
     */
    number: string;
    /**
     * The date of issue, as an ISO 8601 date: `2026-11-30`. Anything after the
     * tenth character (a time) is ignored. Required: the deriver throws when it
     * is missing or does not start with YYYY-MM-DD.
     */
    issueDate: string;
    /**
     * When the goods or services were supplied (the tax point), ISO 8601.
     * Missing or empty means the same day as `issueDate`. Printed only when it
     * is a different day, except in Germany, where a supply line is always
     * printed. Ignored when `supplyPeriod` is given. Throws when present but
     * not an ISO date.
     */
    supplyDate?: string;
    /**
     * A supply that ran over a period (a subscription, a month of service),
     * ISO 8601 at both ends. Printed whenever given, in place of `supplyDate`.
     * Throws when either end is missing or not a date, or when `to` is before
     * `from`.
     */
    supplyPeriod?: { from: string; to: string };
    /**
     * What the document is. `proforma` is a request for payment, not an
     * invoice. `self-billed` is issued by the customer on the supplier's
     * behalf. Required: anything else throws, because a credit note headed
     * "VAT invoice" is worse than no document.
     */
    kind: "invoice" | "proforma" | "self-billed";
    /**
     * Whose rules apply. `US` has no VAT, so every VAT rule is skipped. `EU`
     * means an EU member state with no rules of its own here: the Directive's
     * defaults. Required: anything else throws.
     */
    jurisdiction: "UK" | "EU" | "DE" | "FR" | "US";
    /**
     * The VAT-inclusive total, in the jurisdiction's currency (GBP for UK,
     * EUR otherwise). Read only to decide whether a simplified invoice is
     * allowed, and never printed. Missing or not a finite number means "not
     * simplified", which is always lawful.
     */
    grossTotal: number;
    /** The customer's purchase order. Printed when given; omitted when missing or empty. */
    purchaseOrder?: string;
  };
  /**
   * The supplier: the party whose VAT this is. On a self-billed invoice that
   * is not the party who sent the document, which is why this is `seller`
   * and not the `sender` the letterhead reads.
   */
  seller: {
    /**
     * Whether the supplier is registered for VAT. Missing means not
     * registered: "Invoice", no VAT number, no VAT wording.
     */
    vatRegistered: boolean;
    /**
     * The VAT number, printed as given (trimmed). Required when
     * `vatRegistered` is true outside the US: the deriver throws rather than
     * head a document "VAT invoice" without one. Ignored when not registered.
     */
    vatNumber?: string;
    /**
     * Whether this seller issues simplified invoices at all. Missing means
     * false: a full VAT invoice is always valid, a simplified one is a
     * permission.
     */
    simplifiedAllowed?: boolean;
  };
}

/**
 * Where the lines sit.
 *
 * - `stacked`: one line under another, each with its label. The default.
 * - `table`: labels in a narrow column, values beside them, so the values
 *   line up down the page.
 * - `split`: what identifies the invoice on the left (number, purchase order,
 *   VAT number), its dates ranged right, the way an invoice head often faces
 *   the letterhead.
 */
export type InvoiceDocumentTypeArrangement = "stacked" | "table" | "split";

/** Props for {@linkcode InvoiceDocumentType}. */
export interface InvoiceDocumentTypeProps {
  /** Where the lines sit. Defaults to `"stacked"`. */
  arrangement?: InvoiceDocumentTypeArrangement;
}

export const InvoiceDocumentType: SectionComponent<InvoiceDocumentTypeProps> = async ({ arrangement }) => {
  const [invoice] = useState((data: InvoiceDocumentTypeData) => data.invoice);
  const [seller] = useState((data: InvoiceDocumentTypeData) => data.seller);
  const type = await useDeriver(invoiceDocumentType, [invoice, seller]);

  const lines: InvoiceDocumentTypeLines = {
    statement: type.statement,
    number: invoice.number,
    issueDate: type.issueDate,
    supplyLine: type.supplyLine,
    supplyLabel: type.supplyLabel,
    supplyValue: type.supplyValue,
    purchaseOrder: invoice.purchaseOrder,
    vatNumber: type.vatNumber,
  };

  return (
    <Section id="invoice-type" title={type.heading}>
      {arrange(arrangement, {
        stacked: () => stacked(lines),
        table: () => table(lines),
        split: () => split(lines),
      })}
    </Section>
  );
};
