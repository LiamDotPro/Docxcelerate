import { deriver } from "docxcelerate/template";

/**
 * The VAT a UK domestic reverse charge would have attracted, in sterling.
 *
 * Sterling whatever the invoice currency, because HMRC requires the VAT on a
 * UK VAT invoice to be stated in pounds. It does not round: the caller passes
 * the figure already rounded under its own VAT rounding policy, and this only
 * formats it. A figure that is not a number throws rather than printing
 * "£NaN" on a legal notice. No placeholder, so previews run it.
 *
 * Installed by `dxcl add tax-treatment-notice` as
 * `derivers/tax-treatment-notice.deriver.ts`. The installer registers only
 * node exports, so add it to `derivers/index.ts` by hand:
 * `import taxTreatmentNoticeVat from "./tax-treatment-notice.deriver.ts";`
 * and `taxTreatmentNoticeVat: taxTreatmentNoticeVat.run,` in the object.
 */
export default deriver({
  name: "taxTreatmentNoticeVat",
  run: (amount: unknown): string => {
    const n = Number(amount);

    if (!Number.isFinite(n)) {
      throw new Error(`taxTreatmentNoticeVat needs a finite amount in pounds sterling; got ${String(amount)}.`);
    }

    return new Intl.NumberFormat("en-GB", { style: "currency", currency: "GBP" }).format(n);
  },
});
