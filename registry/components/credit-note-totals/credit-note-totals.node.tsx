import { Cell, Paragraph, Row, Section, type Section as SectionComponent, Table, useDeriver, useState } from "docxcelerate/template";
import creditNoteTotals from "../derivers/credit-note-totals.deriver.ts";

/**
 * What a credit note credits, by VAT band, worked from the invoice it undoes.
 *
 * Credits are worked from the original invoice, never from today's settings.
 * Each credited line keeps the VAT rate the invoice charged, and a
 * foreign-currency invoice keeps its sterling exchange rate, so a rate change
 * since cannot credit the wrong VAT. There is deliberately no input for a
 * current rate.
 *
 * A full cancellation copies the original band figures exactly rather than
 * recomputing them, so the two documents net to zero with no penny residue.
 * A partial return credits quantity times the original unit price, each line
 * rounded once, with VAT rounded per band the way the original was. A goodwill
 * amount is spread across the original VAT bands pro rata by largest
 * remainder, so the figures add up to the penny and no band is credited more
 * VAT than it charged.
 *
 * Earlier credit notes are subtracted first, and a credit larger than what
 * remains in any band fails the document, naming the band, rather than
 * issuing a credit note for money that was never charged.
 *
 * Amounts print positive under a "Credit" heading, never as negative invoice
 * lines. That is what Peppol BIS document type 381 expects: the document type
 * carries the sign.
 *
 * All of it is one deriver, because arithmetic on request data throws at
 * publish. The node does none: it prints the deriver's strings and loops over
 * its rows. The mode is not a branch, it only changes the derived `basis`
 * sentence, so four conditionals and two loops are published.
 *
 * Every line is a figure the reader checks against the original invoice, so
 * they stand in right-aligned columns and there is no `arrangement` prop.
 *
 * Installed by `dxcl add credit-note-totals`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface CreditNoteTotalsData {
  creditNote: {
    /**
     * The invoice being credited, exactly as it was issued. Every rate, price
     * and exchange rate is read from here. There is deliberately no field for
     * a current rate. Required: missing throws.
     */
    original: {
      /**
       * The original invoice number, printed as given (trimmed). Required:
       * missing or blank throws. A credit note must identify what it credits.
       */
      number: string;
      /**
       * The original invoice's date of issue, ISO 8601 (`2026-09-30`). Anything
       * after the tenth character is ignored. Printed, formatted en-GB long, in UTC.
       * Required: missing, or not starting with YYYY-MM-DD, throws.
       */
      issueDate: string;
      /**
       * ISO 4217 code of the original invoice: "GBP", "EUR", "JPY". The credit
       * is in the same currency. The exponent (decimal places) comes from Intl.
       * Required: missing or a code Intl rejects throws, naming
       * `creditNote.original.currency`.
       */
      currency: string;
      /**
       * Pounds sterling per ONE unit of the invoice currency, as used on the
       * original invoice: 0.8567 for EUR. This is NOT the "units per £1" figure
       * HMRC publishes (1.1673). Up to 6 decimal places, honoured exactly.
       * Given and currency is not GBP: a "VAT credited in sterling" row prints.
       * Missing: no sterling row. The invoice did not carry UK VAT.
       * Ignored when currency is GBP. Zero, negative or not finite throws.
       */
      fxToGBP?: number;
      /**
       * The total VAT in sterling exactly as printed on the original invoice.
       * Read only for a full cancellation with no earlier credits, where it is
       * copied rather than recomputed. Missing: computed from `fxToGBP`.
       * Ignored when `fxToGBP` is not in play.
       */
      vatGBP?: number;
      /**
       * The original invoice's VAT analysis: one entry per rate, in the order
       * that invoice printed them. The credit note prints bands in this order.
       * Required and non-empty: missing or empty throws.
       */
      bands: Array<{
        /**
         * VAT rate as a percentage: 20, 5, 12.5, 0. Up to 2 decimal places.
         * `null` is exempt or outside the scope of VAT. 0 is zero-rated, a
         * different treatment, and prints as "0%". Two bands with the same
         * rate throw. `undefined`, negative or not finite throws.
         */
        vatRate: number | null;
        /**
         * Printed instead of the rate for a `null` band: "Exempt", "Outside the
         * scope of VAT". Missing on a null band prints "No VAT". Ignored on a
         * rated band, which is always labelled "20%".
         */
        label?: string;
        /**
         * The band's net as printed on the original, in major units. No more
         * decimals than the currency's exponent: more throws. Negative or not
         * finite throws.
         */
        net: number;
        /**
         * The band's VAT as printed on the original, same rules as `net`.
         * Non-zero on a `null` band throws.
         */
        vat: number;
      }>;
      /**
       * The original invoice's lines. Required only for `credit.mode: "lines"`,
       * where missing throws. Ignored otherwise.
       */
      lines?: Array<{
        /** The line's identifier, unique on the invoice. Duplicates throw. */
        id: string;
        /** What was supplied, printed on the credited line. Wraps; never truncated. */
        description: string;
        /**
         * How many were invoiced. Up to 4 decimal places. Caps what one credit
         * may return of this line. Not finite or ≤ 0 throws.
         */
        qty: number;
        /**
         * Net unit price in major units, as invoiced. Up to 4 decimal places,
         * honoured exactly. Negative or not finite throws.
         */
        unitNet: number;
        /**
         * The rate this line was invoiced at. It must equal some
         * `bands[].vatRate`; one that does not throws.
         */
        vatRate: number | null;
      }>;
    };
    /**
     * Credit notes already issued against this invoice, in any order. Each
     * one's band figures are subtracted from the original before this credit
     * is checked. Missing or empty: nothing has been credited yet.
     */
    previousCredits?: Array<{
      /**
       * The figures that credit note credited, one per rate it touched. A rate
       * not on the original throws. Same number rules as `original.bands`.
       */
      bands: Array<{ vatRate: number | null; net: number; vat: number }>;
    }>;
    /** What this credit note credits. Required: missing throws. */
    credit: {
      /**
       * `"full"`: cancels everything still creditable, copying the figures.
       * `"lines"`: returns some quantity of named original lines.
       * `"amount"`: a lump sum spread across the original bands.
       * Anything else throws.
       */
      mode: "full" | "lines" | "amount";
      /**
       * For `"lines"`: what is returned, in the order it should print. Missing
       * or empty there throws. Ignored for other modes.
       */
      lines?: Array<{
        /** An `original.lines[].id`. Unknown throws; the same id twice throws. */
        id: string;
        /**
         * How many are credited. Up to 4 decimal places. ≤ 0, not finite, or
         * more than the original line's `qty` throws.
         */
        qty: number;
      }>;
      /**
       * For `"amount"`: the VAT-INCLUSIVE sum credited, in major units of the
       * invoice currency. No more decimals than the exponent. Missing, ≤ 0, or
       * more than the total still creditable (gross) throws. Ignored for
       * other modes.
       */
      amount?: number;
    };
  };
}

/**
 * Prints what a credit note credits: how it was worked out, the credited lines
 * when lines were returned, the VAT analysis when more than one band is
 * credited, and the net, VAT and total.
 */
export const CreditNoteTotals: SectionComponent = async () => {
  const [note] = useState((d: CreditNoteTotalsData) => d.creditNote);
  const totals = await useDeriver(creditNoteTotals, [note]);

  return (
    <Section id="credit-note-totals" title="Credit">
      <Paragraph id="credit-note-totals-basis">{totals.basis}</Paragraph>
      <Paragraph id="credit-note-totals-original">{totals.original}</Paragraph>
      {totals.previouslyCredited && (
        <Paragraph id="credit-note-totals-previous">{totals.previouslyCredited}</Paragraph>
      )}
      {totals.hasLines && (
        <Table
          id="credit-note-totals-lines"
          columns={[
            { width: "auto" },
            { width: 22, align: "right" },
            { width: 26, align: "right" },
            { width: 18, align: "right" },
            { width: 28, align: "right" },
          ]}
        >
          <Row id="credit-note-totals-lines-head" header>
            <Cell>Description</Cell>
            <Cell>Qty</Cell>
            <Cell>Unit price</Cell>
            <Cell>VAT</Cell>
            <Cell>Credited</Cell>
          </Row>
          {totals.lines.map((line) => (
            <Row id="credit-note-totals-line">
              <Cell>{line.description}</Cell>
              <Cell>{line.quantity}</Cell>
              <Cell>{line.unitPrice}</Cell>
              <Cell>{line.rate}</Cell>
              <Cell>{line.amount}</Cell>
            </Row>
          ))}
        </Table>
      )}
      {totals.multipleBands && (
        <Table
          id="credit-note-totals-analysis"
          columns={[{ width: "auto" }, { width: 34, align: "right" }, { width: 34, align: "right" }]}
        >
          <Row id="credit-note-totals-analysis-head" header>
            <Cell>Rate</Cell>
            <Cell>Net</Cell>
            <Cell>VAT</Cell>
          </Row>
          {totals.bands.map((band) => (
            <Row id="credit-note-totals-band">
              <Cell>{band.label}</Cell>
              <Cell>{band.net}</Cell>
              <Cell>{band.vat}</Cell>
            </Row>
          ))}
        </Table>
      )}
      <Table id="credit-note-totals-summary" columns={[{ width: "auto" }, { width: 40, align: "right" }]}>
        <Row id="credit-note-totals-net">
          <Cell>Net credited</Cell>
          <Cell>{totals.net}</Cell>
        </Row>
        <Row id="credit-note-totals-vat">
          <Cell>VAT credited</Cell>
          <Cell>{totals.vat}</Cell>
        </Row>
        {totals.vatGBP && (
          <Row id="credit-note-totals-vat-gbp">
            <Cell>VAT credited in sterling</Cell>
            <Cell>{totals.vatGBP}</Cell>
          </Row>
        )}
        <Row id="credit-note-totals-total">
          <Cell>Total credited</Cell>
          <Cell>{totals.total}</Cell>
        </Row>
      </Table>
    </Section>
  );
};
