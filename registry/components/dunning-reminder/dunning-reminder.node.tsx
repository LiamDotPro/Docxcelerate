import {
  arrange,
  Paragraph,
  type Paragraph as ParagraphComponent,
  type Nodes,
  Section,
  useDeriver,
  useSetPlaceholders,
  useSetPrompts,
  useState,
} from "docxcelerate/template";
import type { DunningReminderLines } from "../arrangements/dunning-reminder/lines.ts";
import { stacked } from "../arrangements/dunning-reminder/stacked.arrangement.tsx";
import { table } from "../arrangements/dunning-reminder/table.arrangement.tsx";
import dunningReminder from "../derivers/dunning-reminder.deriver.ts";

/**
 * The body of a payment reminder. It works out which overdue invoices this
 * letter covers, which reminder level it is, its title, its deadline and what
 * it may say will happen next, then prints those lines and nothing else.
 *
 * Written by hand, reminders escalate on days overdue alone, so an invoice
 * jumps from first reminder to final notice, or a daily run sends the same
 * final notice every day. They chase disputed invoices, invoices on a payment
 * plan and balances already paid to zero. They set deadlines shorter than the
 * law allows, and threaten court action the business never takes, which is
 * itself unlawful (UK CPUTR misleading actions; US FDCPA section 807(5)).
 *
 * Decisions already made for you:
 *
 * - An invoice is chased only if it is past due by the grace period, is not
 *   disputed, on hold or on a payment plan, still has an open amount, and
 *   enough time has passed since its last reminder.
 * - An invoice moves up one level per reminder, never more.
 * - The deadline is never shorter than the named legal minimum and never falls
 *   on a weekend.
 * - The final level threatens only the consequences the caller has marked as
 *   actually intended. With none marked, it threatens nothing.
 * - Every level says which date payments were checked up to, because payments
 *   cross in the post.
 * - If nothing qualifies, it prints nothing. It never prints a reminder for
 *   zero invoices.
 *
 * All of it is decided in one deriver, because every step does date
 * arithmetic, filtering, sorting or summing on request data, and publishing
 * refuses all of those.
 *
 * There are eight fixed invoice slots rather than a `.map()`. The resolver
 * reads a published loop's collection from request data and ignores the
 * derived scope, so a `.map()` over derived invoices would publish and then
 * resolve to nothing; conditions and `{{derived.…}}` tokens do honour it.
 *
 * The node publishes 12 branches with the fixed opening and 15 with the
 * generated one, against the default `branchLimit` of 32.
 *
 * The caller must record `lastLevel` and `lastReminderDate` for each invoice
 * after sending. A document cannot write back. The deriver's `sent` field lists
 * them for a caller that runs `run` itself.
 *
 * Installed by `dxcl add dunning-reminder`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface DunningReminderData {
  reminder: {
    /**
     * The date this reminder is sent, ISO 8601 (`2026-11-20`). Days overdue,
     * the interval since the last reminder and the deadline are all counted
     * from it, in whole UTC days. Anything after the tenth character is
     * ignored. Required: the deriver throws when it is missing or not YYYY-MM-DD.
     */
    sendDate: string;
    /**
     * The last day whose payments are reflected in `openAmount`, ISO 8601.
     * Printed in the "if you have paid since" line at every level. Required:
     * throws when missing, not a date, or after `sendDate`.
     */
    paymentsCheckedTo: string;
    /**
     * ISO 4217 code every `openAmount` is in. It also sets the minor units
     * amounts are rounded to (2 for GBP, 0 for JPY). Required: throws when
     * missing or not a code `Intl.NumberFormat` accepts.
     */
    currency: string;
    /**
     * Days past the due date before an invoice is chased at all. Missing
     * means 7. An invoice is never chased on or before its due date, even at
     * 0. Throws when present and not a non-negative integer.
     */
    graceDays?: number;
    /**
     * Minimum days between two reminders for the same invoice. Missing
     * means 14. An invoice reminded more recently than this is left out of
     * this letter altogether, so a daily run neither escalates it nor sends
     * the same level again. Throws when present and not a positive integer.
     */
    minIntervalDays?: number;
    /**
     * Days from `sendDate` to the payment deadline at each level. Missing
     * fields keep their defaults: first 14, second 10, final 7. A legal floor
     * can push the deadline later than this, never earlier. Throws when a
     * given field is not a positive integer.
     */
    deadlineDays?: { first?: number; second?: number; final?: number };
    /**
     * A statutory minimum the deadline may not fall before. Missing means
     * no floor.
     * - `NL-consumer`: 14 days after the day of receipt (BW 6:96(6)).
     * - `BE-consumer`: 14 days counted from the third working day after
     *   sending on paper, or from the day after sending electronically
     *   (WER art. XIX.2).
     * - `UK-pre-action`: 30 days from the date of the letter (Pre-Action
     *   Protocol for Debt Claims).
     * Throws on any other value.
     */
    legalFloor?: "NL-consumer" | "BE-consumer" | "UK-pre-action";
    /**
     * How the reminder is delivered. Read only for the NL and BE floors.
     * Missing means `"post"`, which gives the later of the two deadlines.
     * Throws on any other value.
     */
    channel?: "post" | "email";
    /**
     * The consequences the business actually intends if the final reminder
     * goes unpaid. Printed only on a final-level letter. Missing or empty
     * prints no consequence at all, because threatening an action that is
     * not intended is a misleading practice. Duplicates are ignored. Throws
     * on any other value.
     */
    intendedConsequences?: Array<"agency" | "court" | "suspension">;
    /**
     * When the customer relationship began, ISO 8601. Used only to tell the
     * generated opening how long the customer has been with you; never
     * printed. Missing leaves that fact out of the prompt. Throws when
     * present and not a date, or after `sendDate`.
     */
    customerSince?: string;
    /**
     * Every open invoice on the account, in any order. The deriver decides
     * which ones this reminder covers. Missing or empty means nothing is
     * chased and the component prints nothing.
     */
    invoices?: Array<{
      /** The invoice number, printed as given. Required: throws when missing or blank. */
      number: string;
      /** The due date, ISO 8601. Required: throws when missing or not a date. */
      dueDate: string;
      /**
       * What is still unpaid, in major units (`1234.56`), net of payments up
       * to `paymentsCheckedTo` and of any credit already allocated to it.
       * Rounded half-up to minor units. Zero or negative after rounding is
       * not chased. Required: throws when not a finite number.
       */
      openAmount: number;
      /** True when the customer has disputed it. Missing means false. A disputed invoice is never chased. */
      disputed?: boolean;
      /**
       * True when collection is paused: on hold, or covered by an active
       * payment arrangement. Missing means false. A held invoice is never chased.
       */
      onHold?: boolean;
      /**
       * The highest reminder level already sent for this invoice: 0 none, 1
       * first, 2 second, 3 final. Missing means 0. An invoice at 3 has had its
       * final reminder and is not chased again; what follows is the action,
       * not another letter. Throws when not an integer from 0 to 3.
       */
      lastLevel?: 0 | 1 | 2 | 3;
      /**
       * When the reminder at `lastLevel` was sent, ISO 8601. Required when
       * `lastLevel` is 1 or more: throws when it is missing then, because
       * without it the interval cannot be checked. Ignored at 0. Throws when
       * not a date, or after `sendDate`.
       */
      lastReminderDate?: string;
    }>;
  };
}

/**
 * Where the invoice lines sit.
 *
 * - `table`: number, due date, days overdue and amount in columns, with a total row. The default.
 * - `stacked`: one sentence per invoice down the text column, then the total.
 */
export type DunningReminderArrangement = "table" | "stacked";

/** Props for {@linkcode DunningReminder}. */
export interface DunningReminderProps {
  /** Where the invoice lines sit. Defaults to `"table"`. */
  arrangement?: DunningReminderArrangement;
  /**
   * Generate the opening paragraph per document instead of printing the
   * fixed one. Missing means false. Previews show the level's placeholder.
   */
  generateOpening?: boolean;
}

// Module level is fine here: only the deriver travels as source text.
const OPENINGS = {
  first: {
    general:
      "Write the opening paragraph of a first payment reminder: one or two sentences, friendly, assuming the payment was simply overlooked.",
    placeholder: "A friendly opening, written per document, that assumes the payment was overlooked.",
  },
  second: {
    general:
      "Write the opening paragraph of a second payment reminder: one or two sentences, firm and courteous, noting that an earlier reminder went unanswered.",
    placeholder: "A firm opening, written per document, noting this is the second reminder.",
  },
  final: {
    general:
      "Write the opening paragraph of a final payment reminder: one or two sentences, serious and courteous, saying plainly that this is the final reminder.",
    placeholder: "A serious opening, written per document, stating this is the final reminder.",
  },
};

/** Its hooks belong to its own instance, so the node's `useDeriver` stays its last hook. */
const Opening: ParagraphComponent<{ level: "first" | "second" | "final"; facts: string }> = ({ level, facts }) => {
  useSetPrompts({
    systemPrompt:
      "You write payment reminders on behalf of the sender, in plain English and the second person. Courteous at every level, never sarcastic.",
    generalPrompt: OPENINGS[level].general,
    // Interpolated, not passed bare: a published fact is a token, and the prompt is trimmed as text.
    infoPrompt: `${facts}`,
    negativePrompt:
      "Do not mention amounts, dates, deadlines, invoice numbers or what will happen if the customer does not pay: the letter states those itself. Do not threaten any action. Do not apologise for writing. Do not invent contact details.",
  });
  useSetPlaceholders(OPENINGS[level].placeholder);

  return <Paragraph id={`dunning-reminder-opening-${level}`} />;
};

/**
 * Prints the overdue invoices a reminder covers, with its request, contact
 * line, any intended consequence and the payments-checked line. Prints nothing
 * when no invoice qualifies.
 */
export const DunningReminder: Nodes<DunningReminderProps> = async ({ arrangement, generateOpening }) => {
  const [reminder] = useState((data: DunningReminderData) => data.reminder);
  const r = await useDeriver(dunningReminder, [reminder]);

  const lines: DunningReminderLines = { invoices: r.invoices, overflow: r.overflow, total: r.total };

  return (
    <>
      {r.outcome === "remind" && (
        <Section id="dunning-reminder" title={r.title}>
          {!generateOpening && <Paragraph id="dunning-reminder-opening">{r.opening}</Paragraph>}
          {generateOpening && r.level === 1 && <Opening level="first" facts={r.promptFacts} />}
          {generateOpening && r.level === 2 && <Opening level="second" facts={r.promptFacts} />}
          {generateOpening && r.level === 3 && <Opening level="final" facts={r.promptFacts} />}
          {arrange(arrangement, { table: () => table(lines), stacked: () => stacked(lines) })}
          <Paragraph id="dunning-reminder-request">{r.request}</Paragraph>
          {r.contact && <Paragraph id="dunning-reminder-contact">{r.contact}</Paragraph>}
          {r.consequence && <Paragraph id="dunning-reminder-consequence">{r.consequence}</Paragraph>}
          <Paragraph id="dunning-reminder-paid-since">{r.paidSince}</Paragraph>
        </Section>
      )}
    </>
  );
};
