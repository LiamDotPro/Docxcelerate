import {
  Cell,
  defineMessages,
  type Nodes,
  Paragraph,
  Row,
  Section,
  Table,
  useDeriver,
  useLocale,
  useMessages,
  useState,
} from "docxcelerate/template";
import finalInvoiceSettlement from "../derivers/final-invoice-settlement.deriver.ts";

/**
 * The final invoice after advances: the whole supply by VAT rate, each advance
 * deducted on its own row, and the balance. A final invoice (UK final invoice,
 * DE Schlussrechnung, FR facture de solde) restates the supply and then takes
 * off what the advance invoices already covered.
 *
 * Decisions already made for you:
 *
 * - Deductions are exact. Each advance's net and VAT are taken as printed on
 *   it, never re-derived from its gross, so the bands reconcile with what was
 *   already declared and VAT is not charged twice.
 * - Only received advances are deducted, except in the UK. DE section 14(5)
 *   UStG nets off collected part-payments, and EU and FR VAT on an advance
 *   arises when it is received, so an advance invoiced but unpaid is listed as
 *   "invoiced, not yet received" and stays in the amount due. In the UK an
 *   issued advance VAT invoice is already a tax point (VATA 1994 s6(4)): it is
 *   deducted whether or not it is paid, and listed as still payable in addition.
 * - A rate change between the deposit and completion follows the jurisdiction.
 *   UK: the deposit keeps its own rate and the rest takes the current one, so
 *   two bands print. DE, FR and EU: the whole supply is taxed at the rate in
 *   force when it is performed (BMF letter of 30 June 2020), the advance is
 *   still deducted at its invoiced VAT, and the re-rating difference prints as
 *   an "of which" line under the balance, never added to it a second time.
 * - Overpaid is not a negative amount due. It prints "Overpaid, to be refunded
 *   or credited" and a note that a credit note is needed.
 * - No VAT means no VAT columns: a small-business seller and the US get a
 *   two-column table.
 *
 * `received` is all or nothing: an advance received in part is not supported.
 * An advance's VAT is not checked against its net and rate, because it is taken
 * as printed on purpose, nor that it belongs to this supply. Retention, Skonto,
 * payment terms and statutory exemption mentions belong to other components.
 *
 * All of it is decided in one deriver, in integer minor units, because every
 * step is arithmetic, sorting or comparison on request data and publishing
 * refuses all of those. The node computes nothing: it reads the deriver's result
 * and branches on it, with every test written inline so a publish build keeps
 * each arm.
 *
 * The rows are fixed slots (four bands, eight advances) rather than a `.map()`.
 * The resolver reads a published loop's collection from request data and
 * ignores the derived scope, so a `.map()` over the deriver's lists would
 * publish and then resolve to no rows; conditions and `{{derived.…}}` tokens do
 * honour the derived scope.
 *
 * The node publishes 30 branches against the default `branchLimit` of 32.
 * A document that combines this with other branching components must raise
 * `branchLimit`.
 *
 * Ships in English, German and French; the document's locale picks one, and the
 * amounts, dates and percentages follow it.
 *
 * Installed by `dxcl add final-invoice-settlement`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface FinalInvoiceSettlementData {
  invoice: {
    /**
     * Whose rules apply. `UK`: an advance keeps the VAT rate of its own tax
     * point, and an advance VAT invoice counts as declared once issued, so it
     * is deducted whether paid or not. `DE`, `FR` and `EU`: the whole supply is
     * taxed at the rate in force when it is performed, an advance invoiced at
     * another rate is re-rated in an "of which" line, and only advances
     * received are deducted. `US`: no VAT columns, only advances received are
     * deducted. Required: anything else throws.
     */
    jurisdiction: "UK" | "EU" | "DE" | "FR" | "US";
    /**
     * ISO 4217 code every amount is in: `"EUR"`, `"GBP"`, `"JPY"`. It sets how
     * many minor units make one major unit. Required: missing, or a code
     * `Intl` rejects, throws.
     */
    currency: string;
    /**
     * The whole supply this final invoice covers, every line at its full net
     * value, not just the part still owed. Grouped into one band per rate.
     * Required: missing or empty throws.
     */
    lines: Array<{
      /**
       * Net value in integer minor units (cents, pence). May be negative, for
       * a discount line, but a rate band that nets below zero throws: that is
       * a credit note, not a final invoice. Not a safe integer throws.
       */
      net: number;
      /**
       * VAT rate in percent, as in force when the supply was performed: `20`,
       * `19`, `5.5`. At least 0 and under 100, at most two decimals. Required
       * when VAT is shown. When it is not (US, small business), missing means
       * 0 and any other value throws.
       */
      vatRate?: number;
    }>;
    /**
     * Every advance (deposit, Abschlagsrechnung, facture d'acompte) invoiced
     * towards this supply, in any order. They are listed by issue date, and
     * advances with the same date keep their input order. Missing or empty
     * means none: nothing is deducted and the balance is the whole supply. UK:
     * list a request for payment that was not a VAT invoice only once it has
     * been received.
     */
    advances?: Array<{
      /** The advance invoice's number, printed as given (trimmed). Required: blank throws, and so does a number used twice. */
      number: string;
      /** Its date of issue, ISO 8601 (`2026-03-31`). Anything after the tenth character is ignored. Required: not YYYY-MM-DD throws. */
      issueDate: string;
      /** The net printed on the advance invoice, integer minor units, above zero. Deducted exactly as given. Zero, negative or not a safe integer throws. */
      net: number;
      /**
       * The VAT printed on the advance invoice, integer minor units, 0 or
       * more. Deducted exactly as given, never recomputed from `net` or from
       * the gross. Required when VAT is shown. When it is not, missing means 0
       * and any other value throws (an exempt seller's advance that showed VAT
       * must be corrected first).
       */
      vat?: number;
      /** The rate printed on the advance invoice, in percent. The same rules as a line's `vatRate`. */
      vatRate?: number;
      /**
       * Whether the advance has been paid in full. Outside the UK only `true`
       * is deducted. `false` is listed as "invoiced, not yet received" and
       * stays in the amount due. Required: anything but a boolean throws,
       * because both guesses misstate the balance.
       */
      received: boolean;
      /**
       * Only for an advance whose `vatRate` is no longer one of the lines'
       * rates (a rate changed in between) on an invoice whose lines carry
       * more than one rate: the line rate the advance was paid towards. When
       * the lines carry a single rate, missing means that rate. Missing where
       * it is needed throws, and so does a value that is not one of the
       * lines' rates. Ignored when the advance's rate is one of the lines'
       * rates.
       */
      supplyRate?: number;
    }>;
  };
  seller: {
    /**
     * `true` when the seller charges no VAT at all: a German section 19
     * Kleinunternehmer, a French franchise en base, or a UK seller who is not
     * registered. Then no VAT column prints on any row. Missing means VAT is
     * shown, so a VAT seller who forgets the flag gets an error for the
     * missing rates rather than an invoice without VAT.
     */
    smallBusinessExempt?: boolean;
  };
}

const messages = defineMessages("en", {
  title: "Final settlement",
  description: "Description",
  net: "Net",
  vat: "VAT",
  gross: "Gross",
  amount: "Amount",
  band: (rate: string) => `Supply at ${rate} VAT`,
  supplyTotal: "Total supply",
  advanceVat: (number: string, date: string, rate: string) =>
    `Less advance invoice ${number} of ${date}, VAT at ${rate}`,
  advancePlain: (number: string, date: string) => `Less advance invoice ${number} of ${date}`,
  due: "Amount due",
  overpaid: "Overpaid, to be refunded or credited",
  correction: "Of which VAT on advances re-rated to the rate at completion",
  pendingOpen: (list: string) =>
    `Invoiced, not yet received: ${list}. These advances are not deducted above and remain part of the amount due.`,
  pendingDeducted: (list: string) =>
    `Invoiced, not yet received: ${list}. These advance invoices are deducted above because their VAT has ` +
    "already been accounted for, and they remain payable in addition to the amount due.",
  creditNote: (amount: string) =>
    `The advances received exceed the final total by ${amount}. The excess will be refunded or credited, and a ` +
    "credit note is needed to correct the VAT already declared on the advances.",
}, {
  de: {
    title: "Schlussabrechnung",
    description: "Bezeichnung",
    net: "Netto",
    vat: "USt.",
    gross: "Brutto",
    amount: "Betrag",
    band: (rate) => `Leistung zu ${rate} USt.`,
    supplyTotal: "Gesamtleistung",
    advanceVat: (number, date, rate) => `Abzüglich Abschlagsrechnung ${number} vom ${date}, USt. ${rate}`,
    advancePlain: (number, date) => `Abzüglich Abschlagsrechnung ${number} vom ${date}`,
    due: "Noch zu zahlen",
    overpaid: "Überzahlung, wird erstattet oder gutgeschrieben",
    correction:
      "Darin enthalten: Korrektur der Umsatzsteuer auf Abschlagszahlungen auf den Steuersatz bei Leistungserbringung",
    pendingOpen: (list) =>
      `In Rechnung gestellt, noch nicht bezahlt: ${list}. Diese Abschlagsrechnungen sind oben nicht abgezogen und ` +
      "im noch zu zahlenden Betrag enthalten.",
    pendingDeducted: (list) =>
      `In Rechnung gestellt, noch nicht bezahlt: ${list}. Diese Abschlagsrechnungen sind oben abgezogen, weil die ` +
      "Umsatzsteuer darauf bereits erklärt ist, und bleiben zusätzlich zum noch zu zahlenden Betrag fällig.",
    creditNote: (amount) =>
      `Die erhaltenen Abschlagszahlungen übersteigen den Gesamtbetrag um ${amount}. Der Überschuss wird erstattet ` +
      "oder gutgeschrieben; zur Berichtigung der bereits erklärten Umsatzsteuer ist eine Gutschrift bzw. " +
      "Rechnungskorrektur erforderlich.",
  },
  fr: {
    title: "Décompte final",
    description: "Désignation",
    net: "HT",
    vat: "TVA",
    gross: "TTC",
    amount: "Montant",
    band: (rate) => `Prestation au taux de ${rate}`,
    supplyTotal: "Total de la prestation",
    advanceVat: (number, date, rate) => `Moins facture d'acompte ${number} du ${date}, TVA ${rate}`,
    advancePlain: (number, date) => `Moins facture d'acompte ${number} du ${date}`,
    due: "Reste à payer",
    overpaid: "Trop-perçu, à rembourser ou à créditer",
    correction: "Dont régularisation de la TVA sur acomptes au taux applicable à l'achèvement",
    pendingOpen: (list) =>
      `Facturé, pas encore encaissé : ${list}. Ces acomptes ne sont pas déduits ci-dessus et restent compris ` +
      "dans le reste à payer.",
    pendingDeducted: (list) =>
      `Facturé, pas encore encaissé : ${list}. Ces factures d'acompte sont déduites ci-dessus car leur TVA a ` +
      "déjà été déclarée, et restent dues en plus du reste à payer.",
    creditNote: (amount) =>
      `Les acomptes encaissés dépassent le total de ${amount}. L'excédent sera remboursé ou crédité, et un avoir ` +
      "est nécessaire pour corriger la TVA déjà déclarée sur les acomptes.",
  },
});

/**
 * Prints the supply, each deducted advance and the balance, then the notes the
 * outcome needs.
 */
export const FinalInvoiceSettlement: Nodes = async () => {
  const t = useMessages(messages);
  const [invoice] = useState((data: FinalInvoiceSettlementData) => data.invoice);
  const [seller] = useState((data: FinalInvoiceSettlementData) => data.seller);
  const locale = useLocale();
  const s = await useDeriver(finalInvoiceSettlement, [invoice, seller, locale]);

  return (
    <Section id="final-settlement" title={t.title}>
      {s.vatShown
        ? (
          <Table
            id="final-settlement-vat"
            columns={[{ width: "auto" }, { width: 32, align: "right" }, { width: 32, align: "right" }, { width: 32, align: "right" }]}
          >
            <Row id="final-settlement-vat-header" header>
              <Cell>{t.description}</Cell>
              <Cell>{t.net}</Cell>
              <Cell>{t.vat}</Cell>
              <Cell>{t.gross}</Cell>
            </Row>
            {s.bands[0].rate && (
              <Row id="final-settlement-band-1">
                <Cell>{t.band(s.bands[0].rate)}</Cell>
                <Cell>{s.bands[0].net}</Cell>
                <Cell>{s.bands[0].vat}</Cell>
                <Cell>{s.bands[0].gross}</Cell>
              </Row>
            )}
            {s.bands[1].rate && (
              <Row id="final-settlement-band-2">
                <Cell>{t.band(s.bands[1].rate)}</Cell>
                <Cell>{s.bands[1].net}</Cell>
                <Cell>{s.bands[1].vat}</Cell>
                <Cell>{s.bands[1].gross}</Cell>
              </Row>
            )}
            {s.bands[2].rate && (
              <Row id="final-settlement-band-3">
                <Cell>{t.band(s.bands[2].rate)}</Cell>
                <Cell>{s.bands[2].net}</Cell>
                <Cell>{s.bands[2].vat}</Cell>
                <Cell>{s.bands[2].gross}</Cell>
              </Row>
            )}
            {s.bands[3].rate && (
              <Row id="final-settlement-band-4">
                <Cell>{t.band(s.bands[3].rate)}</Cell>
                <Cell>{s.bands[3].net}</Cell>
                <Cell>{s.bands[3].vat}</Cell>
                <Cell>{s.bands[3].gross}</Cell>
              </Row>
            )}
            <Row id="final-settlement-vat-supply">
              <Cell>{t.supplyTotal}</Cell>
              <Cell>{s.supply.net}</Cell>
              <Cell>{s.supply.vat}</Cell>
              <Cell>{s.supply.gross}</Cell>
            </Row>
            {s.deducted[0].number && (
              <Row id="final-settlement-vat-advance-1">
                <Cell>{t.advanceVat(s.deducted[0].number, s.deducted[0].date, s.deducted[0].rate)}</Cell>
                <Cell>{s.deducted[0].net}</Cell>
                <Cell>{s.deducted[0].vat}</Cell>
                <Cell>{s.deducted[0].gross}</Cell>
              </Row>
            )}
            {s.deducted[1].number && (
              <Row id="final-settlement-vat-advance-2">
                <Cell>{t.advanceVat(s.deducted[1].number, s.deducted[1].date, s.deducted[1].rate)}</Cell>
                <Cell>{s.deducted[1].net}</Cell>
                <Cell>{s.deducted[1].vat}</Cell>
                <Cell>{s.deducted[1].gross}</Cell>
              </Row>
            )}
            {s.deducted[2].number && (
              <Row id="final-settlement-vat-advance-3">
                <Cell>{t.advanceVat(s.deducted[2].number, s.deducted[2].date, s.deducted[2].rate)}</Cell>
                <Cell>{s.deducted[2].net}</Cell>
                <Cell>{s.deducted[2].vat}</Cell>
                <Cell>{s.deducted[2].gross}</Cell>
              </Row>
            )}
            {s.deducted[3].number && (
              <Row id="final-settlement-vat-advance-4">
                <Cell>{t.advanceVat(s.deducted[3].number, s.deducted[3].date, s.deducted[3].rate)}</Cell>
                <Cell>{s.deducted[3].net}</Cell>
                <Cell>{s.deducted[3].vat}</Cell>
                <Cell>{s.deducted[3].gross}</Cell>
              </Row>
            )}
            {s.deducted[4].number && (
              <Row id="final-settlement-vat-advance-5">
                <Cell>{t.advanceVat(s.deducted[4].number, s.deducted[4].date, s.deducted[4].rate)}</Cell>
                <Cell>{s.deducted[4].net}</Cell>
                <Cell>{s.deducted[4].vat}</Cell>
                <Cell>{s.deducted[4].gross}</Cell>
              </Row>
            )}
            {s.deducted[5].number && (
              <Row id="final-settlement-vat-advance-6">
                <Cell>{t.advanceVat(s.deducted[5].number, s.deducted[5].date, s.deducted[5].rate)}</Cell>
                <Cell>{s.deducted[5].net}</Cell>
                <Cell>{s.deducted[5].vat}</Cell>
                <Cell>{s.deducted[5].gross}</Cell>
              </Row>
            )}
            {s.deducted[6].number && (
              <Row id="final-settlement-vat-advance-7">
                <Cell>{t.advanceVat(s.deducted[6].number, s.deducted[6].date, s.deducted[6].rate)}</Cell>
                <Cell>{s.deducted[6].net}</Cell>
                <Cell>{s.deducted[6].vat}</Cell>
                <Cell>{s.deducted[6].gross}</Cell>
              </Row>
            )}
            {s.deducted[7].number && (
              <Row id="final-settlement-vat-advance-8">
                <Cell>{t.advanceVat(s.deducted[7].number, s.deducted[7].date, s.deducted[7].rate)}</Cell>
                <Cell>{s.deducted[7].net}</Cell>
                <Cell>{s.deducted[7].vat}</Cell>
                <Cell>{s.deducted[7].gross}</Cell>
              </Row>
            )}
            {s.overpaid
              ? (
                <Row id="final-settlement-vat-overpaid">
                  <Cell>{t.overpaid}</Cell>
                  <Cell>{s.due.net}</Cell>
                  <Cell>{s.due.vat}</Cell>
                  <Cell>{s.due.gross}</Cell>
                </Row>
              )
              : (
                <Row id="final-settlement-vat-due">
                  <Cell>{t.due}</Cell>
                  <Cell>{s.due.net}</Cell>
                  <Cell>{s.due.vat}</Cell>
                  <Cell>{s.due.gross}</Cell>
                </Row>
              )}
            {s.correction && (
              <Row id="final-settlement-vat-correction">
                <Cell>{t.correction}</Cell>
                <Cell></Cell>
                <Cell>{s.correction}</Cell>
                <Cell></Cell>
              </Row>
            )}
          </Table>
        )
        : (
          <Table id="final-settlement-plain" columns={[{ width: "auto" }, { width: 36, align: "right" }]}>
            <Row id="final-settlement-plain-header" header>
              <Cell>{t.description}</Cell>
              <Cell>{t.amount}</Cell>
            </Row>
            <Row id="final-settlement-plain-supply">
              <Cell>{t.supplyTotal}</Cell>
              <Cell>{s.supply.net}</Cell>
            </Row>
            {s.deducted[0].number && (
              <Row id="final-settlement-plain-advance-1">
                <Cell>{t.advancePlain(s.deducted[0].number, s.deducted[0].date)}</Cell>
                <Cell>{s.deducted[0].net}</Cell>
              </Row>
            )}
            {s.deducted[1].number && (
              <Row id="final-settlement-plain-advance-2">
                <Cell>{t.advancePlain(s.deducted[1].number, s.deducted[1].date)}</Cell>
                <Cell>{s.deducted[1].net}</Cell>
              </Row>
            )}
            {s.deducted[2].number && (
              <Row id="final-settlement-plain-advance-3">
                <Cell>{t.advancePlain(s.deducted[2].number, s.deducted[2].date)}</Cell>
                <Cell>{s.deducted[2].net}</Cell>
              </Row>
            )}
            {s.deducted[3].number && (
              <Row id="final-settlement-plain-advance-4">
                <Cell>{t.advancePlain(s.deducted[3].number, s.deducted[3].date)}</Cell>
                <Cell>{s.deducted[3].net}</Cell>
              </Row>
            )}
            {s.deducted[4].number && (
              <Row id="final-settlement-plain-advance-5">
                <Cell>{t.advancePlain(s.deducted[4].number, s.deducted[4].date)}</Cell>
                <Cell>{s.deducted[4].net}</Cell>
              </Row>
            )}
            {s.deducted[5].number && (
              <Row id="final-settlement-plain-advance-6">
                <Cell>{t.advancePlain(s.deducted[5].number, s.deducted[5].date)}</Cell>
                <Cell>{s.deducted[5].net}</Cell>
              </Row>
            )}
            {s.deducted[6].number && (
              <Row id="final-settlement-plain-advance-7">
                <Cell>{t.advancePlain(s.deducted[6].number, s.deducted[6].date)}</Cell>
                <Cell>{s.deducted[6].net}</Cell>
              </Row>
            )}
            {s.deducted[7].number && (
              <Row id="final-settlement-plain-advance-8">
                <Cell>{t.advancePlain(s.deducted[7].number, s.deducted[7].date)}</Cell>
                <Cell>{s.deducted[7].net}</Cell>
              </Row>
            )}
            {s.overpaid
              ? (
                <Row id="final-settlement-plain-overpaid">
                  <Cell>{t.overpaid}</Cell>
                  <Cell>{s.due.net}</Cell>
                </Row>
              )
              : (
                <Row id="final-settlement-plain-due">
                  <Cell>{t.due}</Cell>
                  <Cell>{s.due.net}</Cell>
                </Row>
              )}
          </Table>
        )}
      {s.pendingDeducted && (
        <Paragraph id="final-settlement-pending-deducted">{t.pendingDeducted(s.pendingDeducted)}</Paragraph>
      )}
      {s.pendingOpen && <Paragraph id="final-settlement-pending-open">{t.pendingOpen(s.pendingOpen)}</Paragraph>}
      {s.overpaid && <Paragraph id="final-settlement-credit-note">{t.creditNote(s.due.gross)}</Paragraph>}
    </Section>
  );
};
