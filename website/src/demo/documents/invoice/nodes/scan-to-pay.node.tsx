import { Image, useDeriver, useState } from "docxcelerate/template";
import { paymentQr } from "../derivers.ts";
import type { InvoiceData } from "../types.ts";

/**
 * The code that opens a transfer with the reference already set.
 *
 * Computed, not composed. It was a `generalPrompt` asking an engine to draw a
 * payment QR, which is the wrong instrument: a QR is a deterministic encoding
 * of a string, and a drawn one does not scan. The design says as much itself —
 * the chip under the card reads "deriver: payment.qr".
 *
 * Two derivations of the same code, because Word will not embed an SVG on its
 * own: the screen takes the vector and the `.docx` takes the raster, which is
 * what `fallbackSrc` is for. The node is no longer dynamic — it has a source
 * and no prompts, which is the honest classification.
 */
export const ScanToPay: Image = async () => {
  const [state] = useState((data: InvoiceData) => ({
    reference: data.reference,
    iban: data.sender.bank.iban,
  }));
  // Both hooks are reached before either is awaited: hooks run in call order,
  // and an await between them would put the second one outside the component.
  const vector = useDeriver(paymentQr, [state.iban, "", state.reference, "svg"]);
  const raster = useDeriver(paymentQr, [state.iban, "", state.reference, "png"]);
  const svg = await vector;
  const png = await raster;

  return (
    <Image
      id="scan-to-pay"
      variant="card"
      src={svg}
      fallbackSrc={png}
      alt={`Scan to pay invoice ${state.reference}`}
      width={108}
      height={108}
    />
  );
};
