/**
 * What the exemption certificate check hands each of its arrangements: the
 * sixteen row slots, the summary and the basis, already decided. An
 * arrangement places these and decides nothing.
 *
 * An unused slot has an empty `description`. Test it in the JSX,
 * `{lines.rows[0].description && …}`, so the test is compiled into a condition
 * the engine settles per document.
 *
 * Installed by `dxcl add exemption-certificate-check`, with the arrangements.
 */
export interface ExemptionCertificateCheckLines {
  /** Exactly 16 slots, in invoice order. */
  rows: Array<{ description: string; amount: string; status: "Exempt" | "Taxable" | ""; reason: string }>;
  /** The exempt and taxable totals as one sentence. */
  summary: string;
  /** The date and the list dates the check was judged against. */
  basis: string;
}
