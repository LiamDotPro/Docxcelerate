import { Image, useState } from "docxcelerate/template";
import type { InvoiceData } from "../types.ts";

/**
 * The code that opens a transfer with the reference already set.
 *
 * A picture nobody can draw at build time: it encodes the account, the amount
 * and the reference, so it belongs to the recipient rather than to the
 * template. The node says what it is and what to show meanwhile, and the
 * engine produces the real one per invoice.
 */
export const ScanToPay: Image = () => {
  const [state] = useState((data: InvoiceData) => ({ reference: data.reference }));

  return (
    <Image
      id="scan-to-pay"
      variant="card"
      alt={`Scan to pay invoice ${state.reference}`}
      width={108}
      height={108}
      generalPrompt={`A payment QR code for invoice ${state.reference}, encoding the ` +
        "account details, the amount due and the reference."}
      placeholder="Scan-to-pay code"
    />
  );
};
