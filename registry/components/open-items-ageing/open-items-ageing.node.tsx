import { arrange, Cell, Paragraph, Row, Section, type Section as SectionComponent, Table, useDeriver, useState } from "docxcelerate/template";
import type { OpenItemsAgeingLines } from "../arrangements/open-items-ageing/lines.ts";
import { across } from "../arrangements/open-items-ageing/across.arrangement.tsx";
import { down } from "../arrangements/open-items-ageing/down.arrangement.tsx";
import openItemsAgeing from "../derivers/open-items-ageing.deriver.ts";

/**
 * An open-item statement: every invoice not yet fully settled, with the amount
 * still open on each, and that open balance aged into buckets. The decisions
 * it makes for the reader are the ones a collections call turns on.
 *
 * Age is counted from the due date, not the invoice date, so an invoice on
 * 30-day terms is not "1–30" the day it is raised. Each bucket includes its top
 * edge: 30 days overdue is 1–30, 31 is 31–60, and an invoice due today is not
 * yet due, which is why the last bucket is "Over 90", not "90+". An invoice
 * with no due date is aged from its date plus `defaultTermsDays`, and the
 * basis sentence says so; with no terms either, the deriver throws rather than
 * guess on a document that tells someone they are overdue.
 *
 * Unallocated credit gets its own line. A payment or credit note not set
 * against an invoice is shown beneath the buckets and never netted into the
 * oldest one, so no bucket goes negative and an old debt does not quietly
 * disappear.
 *
 * Disputed invoices are listed and marked, and aged in their own bucket, so
 * they do not inflate the Over 90 figure that drives escalation.
 *
 * Bad data throws. An over-allocated invoice or credit fails the document
 * rather than printing a negative open amount. The six buckets always sum to
 * the open balance, in integer minor units, and their shares (largest
 * remainder) to exactly 100%.
 *
 * Nothing overdue collapses the strip: when nothing undisputed is overdue it
 * becomes one sentence.
 *
 * Arrangement: only the ageing strip is placed by the `arrangement` prop. The
 * rows, the totals and every sentence are fixed here.
 *
 * Every figure is computed in one deriver, because all of it is arithmetic,
 * sorting and date maths on request data. The rows are a `.map()` over the
 * deriver's result, which publishes as a loop over `derived`. The node
 * branches only on the deriver's booleans and strings, about nine published
 * branches against the default `branchLimit` of 32.
 *
 * Installed by `dxcl add open-items-ageing`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface OpenItemsAgeingData {
  account: {
    /**
     * ISO 4217 code every amount is in: `"GBP"`, `"EUR"`, `"JPY"`. It also sets
     * the minor unit each amount is rounded to on the way in (pence for GBP,
     * none for JPY). Required: the deriver throws when it is missing or not a
     * code `Intl` knows.
     */
    currency: string;
    /**
     * The day the statement is drawn up, ISO 8601: `2026-11-30`. Every age is
     * counted to this day. Anything after the tenth character (a time) is
     * ignored. Required: throws when missing or not YYYY-MM-DD.
     */
    statementDate: string;
    /**
     * Days after its invoice date that an invoice with no `dueDate` falls due.
     * A whole number, 0 or more. Missing is fine while every open invoice has
     * a `dueDate`; the deriver throws, naming the invoice, when an open
     * invoice has no `dueDate` and this is missing, rather than guess terms on
     * a document that tells someone they are overdue. Throws when present but
     * not a non-negative integer.
     */
    defaultTermsDays?: number;
    /**
     * Every invoice on the account, including any that a credit's allocations
     * name, settled ones too: those are left out of the listing but have to be
     * here for an allocation to find. Missing or empty means none.
     */
    invoices?: Array<{
      /** The invoice number, printed as given. Required, non-blank, unique: a duplicate throws, because an allocation could not say which one it meant. */
      number: string;
      /** Invoice date, ISO 8601. Required: throws when missing or not a date. */
      date: string;
      /** Due date, ISO 8601. Missing means `date` plus `account.defaultTermsDays`. Throws when present but not a date. */
      dueDate?: string;
      /**
       * Gross amount in major units: `1186420.5`. Rounded once to the
       * currency's minor unit. Zero means settled and is left out. Negative
       * throws: a credit note belongs in `credits`. Not a finite number throws.
       */
      amount: number;
      /** The customer disputes it. Listed, marked "In dispute", and aged in the In dispute bucket whatever its age. Missing means false. */
      disputed?: boolean;
    }>;
    /** Payments and credit notes received. Missing or empty means none. */
    credits?: Array<{
      /** What it is, printed in the unallocated-credit sentence. Anything else throws. */
      type: "payment" | "credit-note";
      /** When it was received or issued, ISO 8601. Printed with the unallocated part. Required. */
      date: string;
      /** Its reference, printed as given. Required, non-blank. */
      number: string;
      /** Its amount in major units, rounded to the minor unit. Must be above zero, or it throws. */
      amount: number;
      /**
       * Which invoices it settles, and how much of each. Missing or empty means
       * none, and the whole credit is unallocated. Each `invoice` must name a
       * number in `invoices` and each `amount` must be above zero, or it
       * throws. Allocations above the credit's own amount throw.
       */
      allocations?: Array<{ invoice: string; amount: number }>;
    }>;
  };
}

/**
 * Where the ageing strip sits.
 *
 * - `across`: the buckets side by side in three rows (name, amount, share), the way a statement usually prints them. The default.
 * - `down`: one row per bucket with its amount and share beside it, for a narrow column or seven-figure amounts.
 */
export type OpenItemsAgeingArrangement = "across" | "down";

/** Props for {@linkcode OpenItemsAgeing}. */
export interface OpenItemsAgeingProps {
  /** Where the ageing strip sits. Defaults to `"across"`. */
  arrangement?: OpenItemsAgeingArrangement;
}

export const OpenItemsAgeing: SectionComponent<OpenItemsAgeingProps> = async ({ arrangement }) => {
  const [account] = useState((data: OpenItemsAgeingData) => data.account);
  const a = await useDeriver(openItemsAgeing, [account]);

  const lines: OpenItemsAgeingLines = { overdue: a.overdue, buckets: a.buckets };

  return (
    <Section id="open-items-ageing" title="Open items">
      {a.empty && <Paragraph id="open-items-ageing-empty">{a.emptyLine}</Paragraph>}
      {!a.empty && <Paragraph id="open-items-ageing-basis">{a.basis}</Paragraph>}
      {a.hasItems && (
        <Table
          id="open-items-ageing-items"
          columns={[
            { width: "auto" },
            { width: 24 },
            { width: 24 },
            { width: 24 },
            { width: 30, align: "right" },
            { width: 30, align: "right" },
          ]}
        >
          <Row id="open-items-ageing-items-header">
            <Cell>Invoice</Cell>
            <Cell>Dated</Cell>
            <Cell>Due</Cell>
            <Cell>Overdue</Cell>
            <Cell>Amount</Cell>
            <Cell>Open</Cell>
          </Row>
          {a.items.map((item) => (
            <Row id="open-items-ageing-item">
              <Cell>{item.number}</Cell>
              <Cell>{item.date}</Cell>
              <Cell>{item.due}</Cell>
              <Cell>{item.age}</Cell>
              <Cell>{item.amount}</Cell>
              <Cell>{item.open}</Cell>
            </Row>
          ))}
        </Table>
      )}
      {arrange(arrangement, {
        across: () => across(lines),
        down: () => down(lines),
      })}
      {a.nothingOverdue && <Paragraph id="open-items-ageing-nothing-overdue">{a.nothingOverdue}</Paragraph>}
      {!a.empty && (
        <Table id="open-items-ageing-totals" columns={[{ width: "auto" }, { width: 40, align: "right" }]}>
          <Row id="open-items-ageing-total-open">
            <Cell>Open invoices</Cell>
            <Cell>{a.totalOpen}</Cell>
          </Row>
          {a.overdue && (
            <Row id="open-items-ageing-total-overdue">
              <Cell>Of which overdue</Cell>
              <Cell>{a.overdueTotal}</Cell>
            </Row>
          )}
          {a.hasUnallocated && (
            <Row id="open-items-ageing-total-unallocated">
              <Cell>Unallocated credit</Cell>
              <Cell>{a.unallocated}</Cell>
            </Row>
          )}
          <Row id="open-items-ageing-total-balance">
            <Cell>{a.balanceLabel}</Cell>
            <Cell>{a.balance}</Cell>
          </Row>
        </Table>
      )}
      {a.hasUnallocated && <Paragraph id="open-items-ageing-unallocated-detail">{a.unallocatedDetail}</Paragraph>}
    </Section>
  );
};
