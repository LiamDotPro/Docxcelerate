import {
  type Nodes,
  Paragraph,
  useDeriver,
  useLocale,
  useSetPlaceholders,
  useSetPrompts,
  useState,
} from "docxcelerate/template";
import invoiceChangeExplainer from "../derivers/invoice-change-explainer.deriver.ts";

/**
 * One generated sentence under an invoice's totals: why this invoice differs
 * from the previous one. Nothing at all, and no heading, when the difference is
 * not worth a sentence.
 *
 * Decisions already made for the reader:
 *
 * - It names the fewest structured causes that cover 90% of the gross movement
 *   (the sum of |effect|), never more than three. A +100 and a -95 are not
 *   reported as "+5, one cause", and it does not list every changed line.
 * - Price and quantity are split so they add up exactly to the line's change:
 *   volume is the change in quantity at the old price, and price is the rest. A
 *   tax change splits the same way into rate and base.
 * - A longer billing period on its own raises a usage bill, so that share is
 *   attributed to period length, not to "usage rose".
 * - A balance brought forward is never part of what changed.
 * - Silent on a first invoice, after a currency change, when either invoice is
 *   a credit note, and when the difference is under the larger of an absolute
 *   floor and a share of the previous total.
 * - The model may restate the chosen causes and their figures and nothing else;
 *   it may not give reasons. The placeholder is the same explanation written
 *   deterministically, so a document with no model still reads correctly.
 *
 * The node computes nothing: comparisons and arithmetic on request data would
 * throw at publish, so all of it is one deriver. `useDeriver` is the last hook
 * because the hook context does not survive its `await`; the prompts and the
 * placeholder therefore carry hand-written `{{derived.invoiceChangeExplainer.*}}`
 * tokens rather than interpolating the stand-in. The test stays inline, never
 * hoisted into a `const`, so it publishes as a branch.
 *
 * Not done: translating the wording (the locale only formats numbers), guessing
 * why anything changed, comparing against more than one invoice, converting
 * currency, or prorating recurring lines.
 *
 * Installed by `dxcl add invoice-change-explainer`, which also installs its deriver.
 */

/** One invoice line, as this component compares it. */
export interface InvoiceChangeLine {
  /**
   * What matches this line to the same line on the other invoice: a SKU, a
   * price id, a plan code. Stable across invoices and never the description,
   * which gets reworded. Required, and unique within one invoice: the deriver
   * throws on a missing, empty or duplicate key.
   */
  key: string;
  /** What the line is called, printed in the sentence as given. Required. */
  description: string;
  /**
   * What sort of line it is, which decides how a difference is described.
   * `usage` lines may be split by period length. `discount` lines start and
   * end rather than appear and vanish. `one-off` lines are charges that are
   * not expected to repeat. `tax` lines split into rate and base. Required:
   * any other value throws.
   */
  kind: "recurring" | "usage" | "one-off" | "discount" | "tax";
  /**
   * How many were charged. Needed on both invoices, together with
   * `unitPrice`, for a price/quantity split. Without it the line's change is
   * one `changed` cause, "X rose from £a to £b". Throws when present but not
   * a finite number.
   */
  quantity?: number;
  /**
   * The price of one unit, in major units. It may have more decimals than the
   * currency has (`0.00015`). The same rules as `quantity`: needed on both
   * sides for a split, and throws when present but not finite.
   */
  unitPrice?: number;
  /**
   * What the line contributes to the total, in major units, signed. A
   * discount is negative. This is the authority: the split is made against
   * it, not against quantity × unitPrice. Required and finite, or the deriver
   * throws.
   */
  amount: number;
  /**
   * For a `tax` line, the rate as a fraction (`0.2` is 20%). Needed and above
   * 0 on both invoices to split a tax change into rate and base. Otherwise the
   * tax line's change is one `changed` cause. Ignored on other kinds.
   */
  taxRate?: number;
}

/** One invoice, as this component compares it. The current and previous invoices share the shape. */
export interface InvoiceChangeSnapshot {
  /**
   * The invoice total as billed, in major units, including tax. Required and
   * finite. `total - (broughtForward ?? 0)` must equal the sum of the line
   * `amount`s once both are converted to minor units, or the deriver throws:
   * an invoice whose lines do not add up cannot be explained.
   */
  total: number;
  /**
   * Any balance carried in from earlier invoices that `total` includes.
   * Subtracted before comparing, and never a cause. Missing means 0.
   */
  broughtForward?: number;
  /**
   * How many days the billing period covered. Only used for the
   * period-length split, which needs it on both invoices. Missing on either
   * side means no split: the whole volume change of a usage line is a
   * `quantity` cause. Throws when present but not a positive integer.
   */
  periodDays?: number;
  /**
   * Missing means `"invoice"`. A credit note on either side makes the
   * component silent: "lower than your credit note" means nothing.
   */
  documentType?: "invoice" | "credit-note";
  /**
   * For `previous` only: its ISO 4217 currency. Missing means the same as
   * `invoiceChange.currency`. A different code makes the component silent.
   * It never converts. Ignored on `current`.
   */
  currency?: string;
  /** Every line on the invoice, in printed order. Required: throws when not an array. Empty is allowed. */
  lines: InvoiceChangeLine[];
}

/** What this component reads. Add these fields to your document data type. */
export interface InvoiceChangeData {
  invoiceChange: {
    /**
     * The current invoice's ISO 4217 currency. It sets the minor units that
     * every figure is rounded to and the symbol printed. Required: the
     * deriver throws when it is missing or `Intl` rejects it.
     */
    currency: string;
    /** The invoice this document is. Required. */
    current: InvoiceChangeSnapshot;
    /**
     * The invoice before it. Missing or `null` means this is the first
     * invoice: the component is silent.
     */
    previous?: InvoiceChangeSnapshot | null;
    /**
     * When a difference is too small to mention. The component is silent when
     * |difference| < max(`absolute`, `pct` × |previous total net of brought
     * forward|). `absolute` is in major units, default `1`. `pct` is a fraction
     * from 0 up to but not including 1, default `0.02`. Missing, or either
     * field missing, uses the default. Throws on a negative `absolute` or a
     * `pct` outside [0, 1).
     */
    threshold?: { absolute?: number; pct?: number };
    /**
     * What the previous invoice is called in the sentence, after "than":
     * `"last month"`, `"your March invoice"`. Missing or empty means
     * `"your last invoice"`.
     */
    periodLabel?: string;
  };
}

/**
 * One dynamic paragraph telling the customer why the total moved since the
 * previous invoice, written from the facts the deriver settles. Renders
 * nothing when the deriver says the change is not worth explaining.
 */
export const InvoiceChangeExplainer: Nodes = async () => {
  const [change] = useState((data: InvoiceChangeData) => data.invoiceChange);
  // A deriver cannot see the build locale, so it is handed over as an input.
  const locale = useLocale();
  useSetPrompts({
    systemPrompt:
      "You write one or two plain sentences on an invoice, to the customer. You explain figures you are given. You never speculate about why anything changed.",
    generalPrompt:
      "Tell the reader their total is {{derived.invoiceChangeExplainer.difference}} {{derived.invoiceChangeExplainer.direction}} than {{derived.invoiceChangeExplainer.periodLabel}}, and why, using only the changes listed. At most two sentences.",
    infoPrompt: "{{derived.invoiceChangeExplainer.facts}}",
    negativePrompt:
      "Do not give a reason for any change beyond what is listed: nothing about activity, growth, seasons, usage habits or the customer's business. Do not mention any change, balance brought forward, payment or figure that is not listed. Do not round, convert or recompute the figures. Do not apologise or sell.",
    examplePrompt: "Your total is £12.40 higher than last month, mostly because API calls rose from 120,000 to 180,000 (+£9.00).",
  });
  useSetPlaceholders("{{derived.invoiceChangeExplainer.sentence}}");
  const explained = await useDeriver(invoiceChangeExplainer, [change, locale]);

  return <>{explained.speak && <Paragraph id="invoice-change-explainer" />}</>;
};
