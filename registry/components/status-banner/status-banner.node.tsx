import { type Nodes, Paragraph, Shape, useState } from "docxcelerate/template";

/**
 * The coloured block at the top of a document saying where things stand.
 *
 * A drawn shape rather than a filled paragraph, and that is the whole reason
 * it exists as a component. A banner has to be the same depth on every
 * document it goes out on — the one that says "Approved" and the one that
 * says "Awaiting signature — 14 days" have to be the same block, or a reader
 * flicking through a stack sees the shorter one as a different kind of notice.
 * A paragraph with a background grows with its words; a `<Shape>` is the size
 * you gave it.
 *
 * What the banner is *about* is deliberately not decided here. It prints the
 * line it is given under one of three tones, so the same component carries an
 * approval, a deadline, a draft watermark or a refusal — a document that wants
 * one of those writes the sentence and picks the tone. A component that knew
 * about invoices would be an invoice component, and there is already one.
 *
 * Three tones, three ids: positive, attention, critical. They are written as
 * three separate shapes rather than one shape with a computed `variant`
 * because a published document carries every arm and lets the engine choose —
 * a variant worked out here would be frozen at publish time and wrong for
 * every recipient after the first.
 *
 * The variants are the theme's to define. This component names what the banner
 * *is* — `bannerPositive`, `bannerAttention`, `bannerCritical` — and never what
 * colour it should be, so restyling the document restyles the banner without
 * the component changing.
 *
 * Installed by `dxcl add status-banner`. It is your copy: change the wording,
 * the height, or the three tones it branches on.
 */

/** What this component reads. Add these fields to your document data type. */
export interface StatusBannerData {
  status: {
    /** The line printed across the banner. */
    label: string;
    /**
     * Which of the three the banner is drawn as.
     *
     * `attention` is the one anything unrecognised falls to, so a tone added
     * upstream draws a banner rather than nothing at all.
     */
    tone?: "positive" | "attention" | "critical";
    /** A second thing on the line — a date, a reference, whatever follows. */
    note?: string;
  };
}

/**
 * How deep the banner is drawn, in points.
 *
 * One number, used by all three tones, because a banner that changed depth
 * with its wording would be three different banners. Change it here and every
 * tone moves together.
 */
const BANNER_HEIGHT_PT = 44;

export const StatusBanner: Nodes = () => {
  const [status] = useState((data: StatusBannerData) => data.status);

  const line = status.note ? `${status.label} — ${status.note}` : status.label;
  const positive = status.tone === "positive";
  const critical = status.tone === "critical";

  return (
    <>
      {positive && (
        <Shape id="status-banner-positive" variant="bannerPositive" height={BANNER_HEIGHT_PT}>
          <Paragraph id="status-banner-positive-line">{line}</Paragraph>
        </Shape>
      )}

      {critical && (
        <Shape id="status-banner-critical" variant="bannerCritical" height={BANNER_HEIGHT_PT}>
          <Paragraph id="status-banner-critical-line">{line}</Paragraph>
        </Shape>
      )}

      {!positive && !critical && (
        <Shape id="status-banner-attention" variant="bannerAttention" height={BANNER_HEIGHT_PT}>
          <Paragraph id="status-banner-attention-line">{line}</Paragraph>
        </Shape>
      )}
    </>
  );
};
