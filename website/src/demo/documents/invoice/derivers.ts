import { deriver } from "docxcelerate";
import type { InvoiceLine } from "./types.ts";

/**
 * What the invoice adds up to.
 *
 * The arithmetic has to be a deriver rather than a `useState` initializer,
 * because the lines do not exist until a request does: `reduce` walks entries,
 * and at publish time there are none to walk. Computed in the initializer this
 * document builds locally and is refused on publish — which makes it a preview,
 * not a template.
 *
 * All three figures come out of one deriver, from one pass over the lines. A
 * total worked out in a second place is a total that can disagree with the
 * first, and on an invoice that is the error nobody forgives.
 *
 * No `placeholder`, so it runs in the preview too: the preview then shows the
 * real arithmetic rather than a stand-in, and the two cannot drift.
 */
export const invoiceTotals = deriver({
  name: "invoiceTotals",
  run: (lines: InvoiceLine[], vatRate: number) => {
    const subtotal = lines.reduce((total, line) => total + line.qty * line.rate, 0);
    const vat = subtotal * vatRate;
    const money = (value: number) =>
      new Intl.NumberFormat("en-GB", { style: "currency", currency: "GBP" }).format(value);

    // Formatted here rather than returned raw and formatted at the call site:
    // reading a derived number to print it is reading request data, and the
    // publish path refuses that exactly as it refuses the arithmetic.
    return {
      subtotal: money(subtotal),
      vat: money(vat),
      total: money(subtotal + vat),
      rate: new Intl.NumberFormat("en-GB", { style: "percent" }).format(vatRate),
    };
  },
});

/**
 * The dates, written the way a reader reads them.
 *
 * `useFormat`'s `date` reads the value while building, which publishing
 * refuses for the same reason as `reduce`: the value belongs to a request that
 * has not been made. Formatting is exactly what a deriver is for.
 */
export const invoiceDates = deriver({
  name: "invoiceDates",
  run: (issueDate: string, dueDate: string) => {
    const written = (value: string) =>
      new Intl.DateTimeFormat("en-GB", { day: "numeric", month: "long", year: "numeric" })
        .format(new Date(value));

    return { issue: written(issueDate), due: written(dueDate) };
  },
});

/**
 * One charge line, in the figures a reader sees.
 *
 * Both the arithmetic and the formatting have to happen here. Inside the loop
 * that walks the lines, `line.qty * line.rate` is a computation on a value
 * that does not exist yet, and `currency(line.rate)` is a reading of one —
 * publishing refuses both, for the same reason.
 */
export const chargeLine = deriver({
  name: "chargeLine",
  run: (qty: number, rate: number) => {
    const money = (value: number) =>
      new Intl.NumberFormat("en-GB", { style: "currency", currency: "GBP" }).format(value);

    return {
      qty: new Intl.NumberFormat("en-GB", { minimumFractionDigits: 1 }).format(qty),
      rate: money(rate),
      amount: money(qty * rate),
    };
  },
});

/**
 * The scan-to-pay code.
 *
 * A QR is a deterministic encoding of a string, not a picture a model should
 * invent — asked for as a prompt it comes back as something that looks like a
 * code and does not scan, which on an invoice is worse than no code at all.
 * So it is derived, from the account details and the reference it encodes.
 *
 * `format` picks the rendering because Word will not embed an SVG on its own:
 * the screen gets the vector and the `.docx` gets the raster, which is what
 * the image's `fallbackSrc` is for.
 *
 * No `placeholder`, so it runs in the preview too — the preview shows a real,
 * scannable code rather than a stand-in for one.
 */
export const paymentQr = deriver({
  name: "paymentQr",
  run: async (iban: string, amount: string, reference: string, format: string) => {
    const QRCode = (await import("qrcode")).default;
    // EPC069-12: the payment URI European banking apps read from a QR.
    const payload = [
      "BCD",
      "002",
      "1",
      "SCT",
      "",
      "",
      iban.replace(/\s+/g, ""),
      amount,
      "",
      reference,
    ].join("\n");
    const options = { margin: 0, width: 288 } as const;

    return format === "svg"
      ? `data:image/svg+xml;utf8,${encodeURIComponent(await QRCode.toString(payload, { ...options, type: "svg" }))}`
      : await QRCode.toDataURL(payload, options);
  },
});
