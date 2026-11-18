/**
 * Choosing between a component's arrangements: the same lines, placed
 * differently.
 *
 * What a component prints and where it puts it are separate decisions. The
 * first is about the data and is made per document; the second is the
 * author's, made once where the component is used, with a prop. Keeping them
 * apart is what lets an arrangement be added without touching a condition.
 *
 * @module
 */

import type { Yield } from "./element.ts";

/**
 * Yields the arrangement a component was asked for.
 *
 * Picked by lookup rather than written as `arrangement === "table" ? … : …`,
 * because the branch compiler publishes a comparison even of two literals —
 * settling it would bake one build's answer into every document. Compared,
 * every arrangement would ship to the engine as a branch and count against
 * the branch limit; looked up, only the chosen one is ever built.
 *
 * The first arrangement written is the default, which is the same rule the
 * registry catalog lists them by.
 *
 * @example A component with two arrangements
 * ```tsx
 * export const Totals: Section<{ arrangement?: "stacked" | "table" }> = ({ arrangement }) => (
 *   <Section id="totals" title="Totals">
 *     {arrange(arrangement, {
 *       stacked: () => <Paragraph id="totals-due">Due: {due}</Paragraph>,
 *       table: () => <Table id="totals-table" columns={[{}, { align: "right" }]}>…</Table>,
 *     })}
 *   </Section>
 * );
 * ```
 *
 * @typeParam K The arrangement names.
 * @param chosen The arrangement asked for, or `undefined` for the first.
 * @param arrangements Each arrangement, by name, as a function yielding it.
 * @returns What the chosen arrangement yields.
 * @throws If `chosen` names no arrangement, so a typo fails the build rather
 * than quietly printing the default.
 */
export function arrange<K extends string>(
  chosen: K | undefined,
  arrangements: Record<K, () => Yield>,
): Yield {
  const names = Object.keys(arrangements) as K[];
  const name = chosen ?? names[0];

  if (name === undefined || !Object.hasOwn(arrangements, name)) {
    throw new Error(
      `Unknown arrangement ${JSON.stringify(chosen)}; expected one of ${names.join(", ")}.`,
    );
  }

  return arrangements[name]();
}
