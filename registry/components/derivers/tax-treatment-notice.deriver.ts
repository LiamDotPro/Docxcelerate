import { deriver } from "docxcelerate/template";

/**
 * The VAT a UK domestic reverse charge would have attracted, in sterling.
 *
 * Sterling whatever the invoice currency, because HMRC requires the VAT on a
 * UK VAT invoice to be stated in pounds. It does not round: the caller passes
 * the figure already rounded under its own VAT rounding policy, and this only
 * formats it. A missing figure (`null` or `undefined`) gives "": the engine
 * runs a node's derivers before any branch, so this runs for every invoice,
 * and the arm's own guard keeps the empty output from printing. Any other
 * value that is not a number throws rather than printing "£NaN" on a legal
 * notice. No placeholder, so previews run it.
 */
export default deriver({
  name: "taxTreatmentNotice",
  run: (amount: number | null | undefined): string => {
    if (amount == null) return "";

    // The type is what the node passes; the data behind it is JSON nobody
    // checked, so a string or NaN still has to be refused here.
    if (!Number.isFinite(amount)) {
      throw new Error(`taxTreatmentNotice needs a finite amount in pounds sterling; got ${String(amount)}.`);
    }

    return new Intl.NumberFormat("en-GB", { style: "currency", currency: "GBP" }).format(amount);
  },
});
