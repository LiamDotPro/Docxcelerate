import { type Nodes, Paragraph, Shape, useFormat, useState } from "docxcelerate/template";

/**
 * The coloured block at the top of a document saying where things stand.
 *
 * A drawn shape rather than a filled paragraph, and that is the whole reason
 * it exists as a component. A banner has to be the same depth on every
 * document it goes out on — the one that says "Paid" and the one that says
 * "Overdue — 14 days" have to be the same block, or a reader flicking through
 * a stack sees the shorter one as a different kind of notice. A paragraph with
 * a background grows with its words; a `<Shape>` is the size you gave it.
 *
 * Three states, three ids: paid, due, or overdue. Branching on the balance
 * rather than printing one sentence that reads oddly at zero is the same
 * decision `payment-summary` makes, and for the same reason — a recipient who
 * owes nothing should not be shown a deadline.
 *
 * The variants are the theme's to define. This component names what the banner
 * *is* — `bannerPaid`, `bannerDue`, `bannerOverdue` — and never what colour it
 * should be, so restyling the document restyles the banner without the
 * component changing. A theme that has not heard of them draws a plain block,
 * which is a banner without an opinion rather than a broken one.
 *
 * Installed by `dxcl add status-banner`. It is your copy: change the wording,
 * the height, or the three states it branches on.
 */

/** What this component reads. Add these fields to your document data type. */
export interface StatusBannerData {
  account: {
    /** What is owed. Zero or less prints the paid banner. */
    balanceDue: number;
    /** When it is due. Anything `Date` can parse. Ignored once it is paid. */
    dueBy?: string | number | Date;
    /** ISO 4217 code. Defaults to GBP. */
    currency?: string;
    /** Days past the due date. Anything above zero prints the overdue banner. */
    daysOverdue?: number;
  };
}

/**
 * How deep the banner is drawn, in points.
 *
 * One number, used by all three states, because a banner that changed depth
 * with its wording would be three different banners. Change it here and every
 * state moves together.
 */
const BANNER_HEIGHT_PT = 44;

export const StatusBanner: Nodes = () => {
  const format = useFormat();
  const [account] = useState((data: StatusBannerData) => data.account);

  const amount = format.currency(
    Math.abs(account.balanceDue),
    account.currency ?? "GBP",
  );
  const overdue = (account.daysOverdue ?? 0) > 0;

  return (
    <>
      {account.balanceDue <= 0 && (
        <Shape id="status-banner-paid" variant="bannerPaid" height={BANNER_HEIGHT_PT}>
          <Paragraph id="status-banner-paid-line">
            Paid in full — nothing to do
          </Paragraph>
        </Shape>
      )}

      {account.balanceDue > 0 && !overdue && (
        <Shape id="status-banner-due" variant="bannerDue" height={BANNER_HEIGHT_PT}>
          <Paragraph id="status-banner-due-line">
            {amount} due{account.dueBy ? ` by ${format.date(account.dueBy)}` : ""}
          </Paragraph>
        </Shape>
      )}

      {account.balanceDue > 0 && overdue && (
        <Shape id="status-banner-overdue" variant="bannerOverdue" height={BANNER_HEIGHT_PT}>
          <Paragraph id="status-banner-overdue-line">
            {amount} overdue — {account.daysOverdue} days past the due date
          </Paragraph>
        </Shape>
      )}
    </>
  );
};
