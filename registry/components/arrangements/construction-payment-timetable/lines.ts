/**
 * What the construction payment timetable hands each of its arrangements:
 * every line it may print, already decided. An arrangement places these and
 * decides nothing, so both print the same lines under the same conditions.
 *
 * A field that is `null` is a line this timetable does not carry. Test it in
 * the JSX, `{lines.noticeLate && …}`, so the test is compiled into a condition
 * the engine settles per document.
 *
 * Installed by `dxcl add construction-payment-timetable`, with the arrangements.
 */
export interface ConstructionPaymentTimetableLines {
  /** False for a residential occupier: only `notApplicable` is printed. */
  applies: boolean;
  /** The one sentence saying the Act does not apply, or null. */
  notApplicable: string | null;
  /** How the due date was reached. */
  basis: string | null;
  /** The payment due date. */
  dueDate: string | null;
  /** The last day for the payer's payment notice. */
  paymentNoticeBy: string | null;
  /** The final date for payment, after any postponement. */
  finalDate: string | null;
  /** The last day for a pay less notice. */
  payLessBy: string | null;
  /** That the payment notice came late and is treated as not given, or null. */
  noticeLate: string | null;
  /** That a valid payment notice was given, or null. */
  noticeGiven: string | null;
  /** That the payee's application stands as the notified sum, or null. */
  applicationStands: string | null;
  /** That the payee's default notice stands, or null. */
  defaultNotice: string | null;
  /** That the payee may still give a default notice, or null. */
  defaultOpen: string | null;
  /** Which days were not counted, and under what. */
  reckoning: string | null;
}
