import { Cell, Paragraph, Row, Section, type Section as SectionComponent, Table, useDeriver, useState } from "docxcelerate/template";
import timesheetFeeSummary from "../derivers/timesheet-fee-summary.deriver.ts";

/**
 * The fee summary of a professional-services invoice, from raw time entries:
 * one row per timekeeper (name, role, hours, rate, amount), then adjustments,
 * then the subtotal. It does not print narratives or an itemised entry list;
 * those belong to a separate detail schedule.
 *
 * Decisions already made for you:
 *
 * - Rounding is per entry, to the billing increment, with a minimum of one
 *   increment. A client's guidelines name one method and the others bill
 *   different hours: three entries of 7 minutes bill 0.60 hours, not 0.40.
 * - Rates follow the entry date. A rate change mid-invoice splits that
 *   timekeeper into two rows, rather than applying the new rate backwards.
 * - A guideline reduction is shown, not hidden. The row keeps the
 *   engagement-letter rate and a separate "Reduction per billing guidelines"
 *   line carries the difference, so the invoice still matches the letter.
 * - Travel is billed at the agreed factor on its own row, no-charge time is
 *   listed at 0.00 rather than dropped, and a write-down is its own line with
 *   its reason.
 * - Figures are exact. Hours always print with two decimals and every row's
 *   hours times its printed rate equals its printed amount. That is why the
 *   billing increment must be a multiple of 3 minutes: only those are exact
 *   in hundredths of an hour, and 10 minutes (0.1666…) would be recomputed
 *   differently by the client's system. All money is integer minor units.
 *
 * All of it is one deriver, because request data is a stand-in at publish time
 * and arithmetic, sorting and filtering on it throw. This node only places
 * what the deriver returns, and loops over its rows and adjustments, so a
 * reduction or write-down costs no branch: it publishes 3 branches.
 *
 * Installed by `dxcl add timesheet-fee-summary`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface TimesheetFeeSummaryData {
  time: {
    /**
     * Every time entry on this invoice, in any order. Each is rounded on its
     * own, then grouped into the timekeeper's rows. An empty array prints "No
     * time was recorded for this period." and no table. Missing, or not an
     * array, throws.
     */
    entries: Array<{
      /**
       * The day the work was done, ISO 8601 (`2026-10-02`). Anything after the
       * tenth character is ignored. Picks the rate row in force that day.
       * Missing or not YYYY-MM-DD throws.
       */
      date: string;
      /**
       * Whose time this is: an `id` in `time.timekeepers`. An id that is not
       * there throws, naming the entry's index.
       */
      timekeeperId: string;
      /**
       * Raw minutes recorded, before rounding. Fractions are allowed. Zero,
       * negative, or not a finite number throws: a correction to time already
       * billed belongs on a credit note, and the minimum-increment rule would
       * turn a 0 into one billed increment.
       */
      minutes: number;
      /**
       * What kind of time this is. Missing means "work". "travel" is billed at
       * `billing.travelFactor` on its own row. "no-charge" is listed with its
       * hours at 0.00 on its own row. Any other value throws.
       */
      kind?: "work" | "travel" | "no-charge";
    }>;
    /**
     * Everyone who may appear, in the order their rows should print (usually
     * seniority). A timekeeper with no entries prints nothing. A duplicate
     * `id` throws.
     */
    timekeepers: Array<{
      /** The key entries refer to. Required, non-empty. */
      id: string;
      /** Printed as given in the first column. Required, non-empty. */
      name: string;
      /** Printed as given in the second column: "Partner", "Trainee Solicitor". Required. */
      role: string;
      /**
       * The standard hourly rate, as dated rows: each applies from its `from`
       * date (ISO 8601, inclusive) until the next row's. In any order. Empty,
       * a duplicate `from`, a rate that is negative or not finite, or a rate
       * with more decimals than the currency has (450.005 GBP) throws. An
       * entry dated before the earliest `from` throws.
       */
      rates: Array<{ from: string; rate: number }>;
      /**
       * The hourly rate the client's billing guidelines approve for this
       * timekeeper, undated. Each row bills at the lower of this and the
       * standard rate. Where this is lower, the row still prints the standard
       * rate and amount, and a reduction line shows the difference. Missing
       * means no approved rate: standard throughout. Same validation as a
       * rate.
       */
      approvedRate?: number;
    }>;
    /** How the engagement bills time. Missing means every default below. */
    billing?: {
      /**
       * The billing increment in whole minutes. Missing means 6 (0.1 hours).
       * Must be a positive integer multiple of 3: only those give hours that
       * are exact in two decimals. 10 minutes is 0.1666… hours, which a
       * client's system would recompute differently. Anything else throws.
       */
      incrementMinutes?: number;
      /**
       * "up" rounds each entry up to the next increment. "nearest" rounds to
       * the nearest, half up. Either way an entry bills at least one
       * increment. Missing means "up", the conventional six-minute rule. Any
       * other value throws.
       */
      rounding?: "up" | "nearest";
      /**
       * The share of the hourly rate travel is billed at: 0.5 for 50%. Must be
       * greater than 0 and at most 1, otherwise it throws. Missing means 1:
       * travel is still on its own row, at the full rate, because inventing a
       * discount the engagement did not agree is as wrong as omitting one.
       */
      travelFactor?: number;
      /**
       * An agreed write-down, in major units, printed as its own negative line
       * with its reason. `amount` 0 prints no line. A negative amount, an
       * empty `reason`, or a write-down larger than the fees after reductions
       * throws: an unexplained or excessive write-down is the line a client
       * queries.
       */
      writeDown?: { amount: number; reason: string };
    };
    /**
     * ISO 4217 code. It sets the number of minor-unit digits for every amount
     * and rate (2 for GBP, 0 for JPY) and is printed in the column headers.
     * Missing means GBP. Not three letters A–Z (after upper-casing) throws.
     */
    currency?: string;
  };
}

export const TimesheetFeeSummary: SectionComponent = async () => {
  const [time] = useState((data: TimesheetFeeSummaryData) => data.time);
  const fees = await useDeriver(timesheetFeeSummary, [time]);

  return (
    <Section id="timesheet-fee-summary" title="Professional fees">
      {fees.hasRows && <Paragraph id="timesheet-fee-basis">{fees.basis}</Paragraph>}
      {fees.hasRows && (
        <Table
          id="timesheet-fee-table"
          columns={[{ width: "auto" }, { width: "auto" }, { width: 18, align: "right" }, { width: 24, align: "right" }, { width: 28, align: "right" }]}
        >
          <Row id="timesheet-fee-header" header>
            <Cell>Timekeeper</Cell>
            <Cell>Role</Cell>
            <Cell>Hours</Cell>
            <Cell>Rate ({fees.currency})</Cell>
            <Cell>Amount ({fees.currency})</Cell>
          </Row>
          {fees.rows.map((row) => (
            <Row id="timesheet-fee-line">
              <Cell>{row.label}</Cell>
              <Cell>{row.role}</Cell>
              <Cell>{row.hours}</Cell>
              <Cell>{row.rate}</Cell>
              <Cell>{row.amount}</Cell>
            </Row>
          ))}
          {fees.adjustments.map((adjustment) => (
            <Row id="timesheet-fee-adjustment">
              <Cell>{adjustment.label}</Cell>
              <Cell>{""}</Cell>
              <Cell>{""}</Cell>
              <Cell>{""}</Cell>
              <Cell>{adjustment.amount}</Cell>
            </Row>
          ))}
          <Row id="timesheet-fee-subtotal">
            <Cell>Subtotal</Cell>
            <Cell>{""}</Cell>
            <Cell>{""}</Cell>
            <Cell>{""}</Cell>
            <Cell>{fees.subtotal}</Cell>
          </Row>
        </Table>
      )}
      {!fees.hasRows && <Paragraph id="timesheet-fee-none">No time was recorded for this period.</Paragraph>}
    </Section>
  );
};
