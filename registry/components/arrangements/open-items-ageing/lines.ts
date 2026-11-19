/**
 * What open items and ageing hands each of its arrangements: the ageing strip,
 * already decided. An arrangement places these and decides nothing, so both
 * print the same figures under the same condition.
 *
 * Installed by `dxcl add open-items-ageing`, with the arrangements.
 */
export interface OpenItemsAgeingLines {
  /** Whether anything undisputed is overdue. When false the strip is not printed at all. */
  overdue: boolean;
  /** Always six buckets, in printed order: not yet due, 1–30, 31–60, 61–90, over 90, in dispute. */
  buckets: Array<{ label: string; amount: string; share: string }>;
}
