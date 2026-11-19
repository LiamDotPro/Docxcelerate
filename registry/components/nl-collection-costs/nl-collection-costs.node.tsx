import {
  arrange,
  type Nodes,
  Paragraph,
  Section,
  useDeriver,
  useState,
} from "docxcelerate/template";
import type { NlCollectionCostsLines } from "../arrangements/nl-collection-costs/lines.ts";
import { stacked } from "../arrangements/nl-collection-costs/stacked.arrangement.tsx";
import { table } from "../arrangements/nl-collection-costs/table.arrangement.tsx";
import nlCollectionCosts from "../derivers/nl-collection-costs.deriver.ts";

/**
 * The statement of Dutch out-of-court collection costs (buitengerechtelijke
 * incassokosten), printed only when the law lets this letter claim them. Each
 * mistake below voids the whole claim in court, so the component decides them.
 *
 * A consumer owes costs only after a veertiendagenbrief that states the exact
 * figure; a business owes them from the day after the due date, with no
 * letter. The fourteen days run from the day after receipt (HR 2016:2704), so
 * the letter prints a payment date, never a period: fifteen days after the
 * expected receipt (post four days, e-mail one), rolled past weekends and
 * Algemene termijnenwet holidays. It can be a day late, never too short.
 *
 * The staffel is applied once to everything the letter names (Besluit BIK
 * art. 2), on integer cents, with statutory interest added first once it has
 * run a year. btw is added only when the creditor is known not to reclaim it:
 * understating the figure keeps the letter valid, overstating does not. A
 * consumer's low instalments are capped (art. 2a). A later letter claims the
 * stated figure only once the earlier deadline has passed and that letter
 * stated one; otherwise no costs are printed at all.
 *
 * A collection agency's Wki registration number prints in every outcome, none
 * included, because the duty is on the letter and not on the costs. That is
 * why this is `Nodes`: the section, and the agency line outside it.
 *
 * All of it is decided in one deriver, because every step is arithmetic, a
 * comparison or a date on request data, which publishing refuses. The node
 * computes nothing and branches with inline tests so a publish build keeps
 * each arm. It publishes 7 branches here (the section, five arms, the agency line) and
 * 2 more in whichever arrangement is used (interest, btw): 9 against the default
 * `branchLimit` of 32.
 *
 * Arrangement: where the figures sit is a separate choice from what the
 * statement says, made with the `arrangement` prop. Both arrangements, in
 * `arrangements/nl-collection-costs/`, print the same lines.
 *
 * Installed by `dxcl add nl-collection-costs`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface NlCollectionCostsData {
  collection: {
    /**
     * Who owes the money. `consumer` is a natural person not acting in a trade
     * or profession: the veertiendagenbrief rules apply. `business` is
     * everyone else: no letter is needed and a contract clause may set the
     * costs. Required: the deriver throws on anything else, because the two
     * rules give different answers for every other field.
     */
    debtorType: "consumer" | "business";
    /**
     * Whether the creditor can reclaim the VAT on the collection costs.
     * `false` (a private person, a VAT-exempt business) adds 21% btw to the
     * costs. Missing means `true`, so no btw is added: on a consumer letter the
     * stated figure must be right, and understating it keeps the letter
     * valid where overstating does not.
     */
    creditorCanReclaimVat?: boolean;
    /**
     * The date this document is sent, ISO 8601 (`2026-11-19`). Anything after
     * the tenth character is ignored. Required: throws when missing or not
     * YYYY-MM-DD. Every invoice must already be overdue on this date.
     */
    sendDate: string;
    /**
     * How this document reaches a consumer. It sets the expected receipt:
     * `post` is `sendDate` plus `postDeliveryDays`, `email` is the next day.
     * Required for a consumer when `previousLetter` is missing (this
     * document is then the veertiendagenbrief), and throws when missing
     * there. Ignored otherwise.
     */
    channel?: "post" | "email";
    /**
     * Calendar days a posted letter is assumed to take. Missing means 4,
     * which covers PostNL's two-delivery-day service whichever weekday the
     * letter is posted. Throws when given and not a non-negative integer.
     */
    postDeliveryDays?: number;
    /**
     * Every invoice the letter claims, at least one. Their principals are
     * added together before the staffel. Required: throws when missing or
     * empty.
     */
    invoices: Array<{
      /** The invoice number, printed as given. Throws when missing or blank. */
      number: string;
      /**
       * What is still owed on it, in euros (`7845.32`). Throws when it is
       * not a finite number above 0: a credit or a paid invoice does not
       * belong in a collection letter.
       */
      principal: number;
      /**
       * When it fell due, ISO 8601. Throws when missing, not a date, or not
       * before `sendDate`. The debtor is in default only from the day after
       * the due date, and a letter sent earlier is invalid.
       */
      dueDate: string;
    }>;
    /**
     * Statutory interest already due, in euros. Added to the principal before
     * the staffel only when `interestRunningSince` is at least one calendar
     * year before `sendDate`. Missing, 0, or interest that has not yet run
     * a year adds nothing. Throws when given and negative or not finite.
     */
    interestAccrued?: number;
    /**
     * The day the interest started to run, ISO 8601. Missing means the
     * interest is never added. Throws when given and not a date.
     */
    interestRunningSince?: string;
    /**
     * For a consumer's instalment debt (art. 2a Besluit BIK). The letter must
     * then name exactly one invoice, the instalment, and the deriver throws
     * otherwise. Ignored for a business debtor. Missing means not an
     * instalment.
     */
    instalment?: {
      /**
       * How many earlier instalments of the same contract were already
       * charged collection costs in the current six-month window. 0 means
       * this is the first. Throws when not a non-negative integer.
       */
      priorCharged: number;
      /**
       * The day the first charged instalment of the window was sent, ISO 8601.
       * The window ends six calendar months later; a `sendDate` on or after
       * that starts a new window, and this instalment is the first again.
       * Required when `priorCharged` is above 0, and throws when missing there.
       */
      windowStart?: string;
    };
    /**
     * For a business debtor: the costs, in euros excluding btw, that the
     * contract's own clause gives. They replace the staffel and the EUR 40
     * per invoice. Missing or 0 means no clause. Ignored for a consumer,
     * because the staffel cannot be contracted out of to a consumer's cost.
     * Throws when negative or not finite.
     */
    contractualCosts?: number;
    /**
     * For a consumer: the veertiendagenbrief that was already sent, when this
     * document is the later letter that claims the costs. Missing means this
     * document IS the veertiendagenbrief. Ignored for a business debtor.
     */
    previousLetter?: {
      /**
       * The payment date that letter stated, ISO 8601. Costs are due only when
       * `sendDate` is after it. Required: throws when missing or not a date.
       */
      deadline: string;
      /**
       * The cost amount that letter stated, in euros including any btw. It
       * is what is claimed now, and it is not recomputed. Missing or 0 means
       * the letter stated no amount, so no costs are due and none are printed.
       * Throws when negative.
       */
      statedCosts?: number;
      /**
       * When the consumer actually received it, ISO 8601, if known. When
       * given, costs are also not due until fourteen days after it have passed
       * (Atw-extended). Throws when given and not a date.
       */
      receivedDate?: string;
    };
    /**
     * Set when a collection agency sends the letter, not the creditor. Both
     * fields are printed, and both are required: throws when either is
     * missing or blank. Missing means the creditor sends it.
     */
    agency?: {
      /** The agency's name as registered. */
      name: string;
      /** Its number in the Wki register (register incassodienstverlening). */
      registrationNumber: string;
    };
  };
}

/**
 * Where the figures sit. `stacked`: labelled lines. `table`: labels left,
 * amounts right-aligned.
 */
export type NlCollectionCostsArrangement = "stacked" | "table";

/** Props for {@linkcode NlCollectionCosts}. */
export interface NlCollectionCostsProps {
  /** Where the figures sit. Defaults to `"stacked"`. */
  arrangement?: NlCollectionCostsArrangement;
}

export const NlCollectionCosts: Nodes<NlCollectionCostsProps> = async ({ arrangement }) => {
  const [collection] = useState((data: NlCollectionCostsData) => data.collection);
  const c = await useDeriver(nlCollectionCosts, [collection]);

  const lines: NlCollectionCostsLines = {
    invoices: c.invoices,
    principal: c.principal,
    interest: c.interest,
    costs: c.costs,
    vat: c.vat,
    total: c.total,
    totalLabel: c.totalLabel,
  };

  return (
    <>
      {c.outcome !== "none" && (
        <Section id="nl-collection-costs" title="Incassokosten">
          {c.outcome === "consumer-notice" && (
            <Paragraph id="nl-collection-costs-notice">
              Wij verzoeken u {c.payNow} voor {c.invoices} uiterlijk op {c.deadline} te betalen. Betaalt u niet
              uiterlijk op die datum, dan bent u daarnaast incassokosten verschuldigd van {c.costsPhrase}, berekend
              volgens het Besluit vergoeding voor buitengerechtelijke incassokosten.
            </Paragraph>
          )}
          {c.outcome === "consumer-due" && (
            <Paragraph id="nl-collection-costs-due">
              De betalingstermijn uit onze eerdere brief is op {c.previousDeadline} verstreken zonder dat {c.payNow}{" "}
              is betaald. U bent daarom de in die brief aangekondigde incassokosten van {c.costsPhrase} verschuldigd.
            </Paragraph>
          )}
          {c.outcome === "business-staffel" && (
            <Paragraph id="nl-collection-costs-business-staffel">
              Omdat {c.invoices} niet op tijd is betaald, bent u incassokosten verschuldigd van {c.costsPhrase},
              berekend volgens het Besluit vergoeding voor buitengerechtelijke incassokosten (artikel 6:96 BW).
            </Paragraph>
          )}
          {c.outcome === "business-fixed" && (
            <Paragraph id="nl-collection-costs-business-fixed">
              Omdat {c.invoices} niet op tijd is betaald, bent u de vaste vergoeding van € 40 per factuur
              verschuldigd (artikel 6:96 lid 4 BW): {c.costsPhrase} voor {c.invoiceCount} facturen.
            </Paragraph>
          )}
          {c.outcome === "business-contract" && (
            <Paragraph id="nl-collection-costs-business-contract">
              Omdat {c.invoices} niet op tijd is betaald, bent u de incassokosten verschuldigd die in de overeenkomst
              zijn afgesproken: {c.costsPhrase}.
            </Paragraph>
          )}
          {arrange(arrangement, { stacked: () => stacked(lines), table: () => table(lines) })}
        </Section>
      )}
      {collection.agency && (
        <Paragraph id="nl-collection-costs-agency">
          {collection.agency.name} is ingeschreven in het register incassodienstverlening onder nummer{" "}
          {collection.agency.registrationNumber}.
        </Paragraph>
      )}
    </>
  );
};
