import { type Nodes, Paragraph, Shape, useState } from "docxcelerate/template";

/**
 * The coloured block at the top of a document saying where things stand.
 *
 * A `<Shape>` rather than a filled paragraph, because a banner has to be the
 * same depth whatever it says — one that grew with its wording would read as a
 * different kind of notice on every document. Three tones, three ids, written
 * as three shapes rather than one with a computed `variant`: a published
 * document carries every arm and lets the engine choose, so a variant worked
 * out here would be frozen at publish time.
 *
 * Installed by `dxcl add status-banner`. It is your copy: change the wording,
 * the height, or the three tones it branches on.
 */

/** What this component reads. Add these fields to your document data type. */
export interface StatusBannerData {
  status: {
    /** The line printed across the banner. */
    label: string;
    /** Which of the three it is drawn as. Anything unrecognised draws `attention`. */
    tone?: "positive" | "attention" | "critical";
    /** A second thing on the line — a date, a reference, whatever follows. */
    note?: string;
  };
}

/** One depth for all three tones, so the wording cannot change the block. */
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
