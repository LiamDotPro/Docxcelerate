import type { DeriverModule } from "../runtime/deriver_module.ts";
import { setPath } from "../runtime/object_path.ts";
import { createDerivedStandIn, expr } from "./publish.ts";
import {
  type ComponentInstance,
  nextCell,
  type PromptDraft,
  requireContext,
  requireInstance,
} from "./context.ts";

/**
 * The hooks a component calls while a build renders it.
 *
 * A build is one pass, so these are not the React hooks they read like: nothing
 * re-renders and nothing is cached between builds. What they do share is the
 * ordering rule — hooks are matched to cells by call order, so a component has
 * to call the same ones in the same order every time.
 *
 * @module
 */

/**
 * A component's initial state, or a function deriving it from the document data.
 *
 * @typeParam TState What the component keeps.
 * @typeParam TData The shape the document reads.
 */
export type StateInitializer<TState, TData> = TState | ((data: TData) => TState);

/**
 * State the component keeps, seeded from the document data once.
 *
 * The initializer is where data enters a component; nothing else reaches for
 * it. Everything after reads state, which is what lets the branch compiler see
 * a component's dependencies instead of guessing at scattered reads.
 *
 * A build is a single pass, so the setter is not a re-render request. It
 * updates the value that this component reads later on, and that anything
 * sharing it through `useShared` will read.
 */
export function useState<TState, TData = unknown>(
  initial: StateInitializer<TState, TData>,
): [TState, (next: TState | ((previous: TState) => TState)) => void] {
  const instance = requireInstance("useState");
  const context = requireContext("useState");
  const cell = nextCell(instance, () =>
    typeof initial === "function"
      ? (initial as (data: TData) => TState)(context.data as TData)
      : initial);

  return [
    cell.value,
    (next) => {
      cell.value = typeof next === "function"
        ? (next as (previous: TState) => TState)(cell.value)
        : next;
    },
  ];
}

/**
 * State shared by everything in one build, keyed by name.
 *
 * Reads see whatever the components rendered before this one left behind —
 * a running total, the facts already stated, the budget already spent. Order is
 * document order, so a later section can react to an earlier one.
 */
export function useShared<TValue>(
  key: string,
  initial: TValue | (() => TValue),
): [TValue, (next: TValue | ((previous: TValue) => TValue)) => void] {
  const context = requireContext("useShared");
  requireInstance("useShared");

  if (!context.shared.has(key)) {
    context.shared.set(key, typeof initial === "function" ? (initial as () => TValue)() : initial);
  }

  return [
    context.shared.get(key) as TValue,
    (next) => {
      const previous = context.shared.get(key) as TValue;
      context.shared.set(
        key,
        typeof next === "function" ? (next as (previous: TValue) => TValue)(previous) : next,
      );
    },
  ];
}

/**
 * There is deliberately no `useMemo`.
 *
 * A build renders each component once and keeps no instances between builds, so
 * a memo could never return a cached value — and caching across builds would be
 * wrong anyway, because each build has different data. Compute in the `useState`
 * initializer instead, which is the one place that runs once by construction.
 */

/** The token budget this build was given, for prompts that need to size themselves. */
export function useAvailableTokens(): number {
  return requireContext("useAvailableTokens").availableTokens;
}

/**
 * Sets the prompts for whatever node this component yields.
 *
 * Calling this is what makes the node dynamic. The returned draft can be spread
 * onto the element instead, when saying so at the element reads better; both
 * routes end in the same place, and props win over the hook so a caller can
 * override what a shared hook set.
 */
export function useSetPrompts(prompts: PromptDraft): PromptDraft {
  const instance = requireInstance("useSetPrompts");
  assign(instance.prompts, prompts);

  return { ...instance.prompts };
}

/**
 * Sets the stand-in text shown wherever a document is previewed rather than
 * written — before an AI client has produced anything.
 */
export function useSetPlaceholders(placeholder: string | { placeholder?: string }): PromptDraft {
  const instance = requireInstance("useSetPlaceholders");

  assign(
    instance.prompts,
    typeof placeholder === "string" ? { placeholder } : placeholder,
  );

  return { ...instance.prompts };
}

/** The stand-in values {@linkcode usePlaceholderData} hands out. */
export interface PlaceholderData {
  /**
   * A person's name.
   *
   * @returns A first and last name.
   */
  name(): string;
  /**
   * A place name.
   *
   * @returns A city.
   */
  city(): string;
  /**
   * A date, formatted for the build's locale.
   *
   * @param offsetDays Days from the fixed base date. Defaults to `0`.
   * @returns The formatted date.
   */
  date(offsetDays?: number): string;
  /**
   * An amount of money, formatted for the build's locale.
   *
   * @param amount The amount. A stable arbitrary one is used when absent.
   * @returns The formatted amount.
   */
  currency(amount?: number): string;
  /**
   * A sentence of filler.
   *
   * @param words How many words. Defaults to `12`.
   * @returns The sentence, capitalised and stopped.
   */
  sentence(words?: number): string;
  /**
   * A paragraph of filler.
   *
   * @param sentences How many sentences. Defaults to `3`.
   * @returns The paragraph.
   */
  paragraph(sentences?: number): string;
  /**
   * One of your own values, chosen the same way every build.
   *
   * @typeParam TValue What the list holds.
   * @param values The values to choose between.
   * @returns One of them.
   */
  pick<TValue>(values: readonly TValue[]): TValue;
}

/**
 * Stand-in values for a preview.
 *
 * Seeded from where the component sits, so the same node shows the same name
 * and the same figures every time it is previewed. A preview that reshuffles
 * itself on each build is one nobody can proofread.
 */
export function usePlaceholderData(): PlaceholderData {
  const instance = requireInstance("usePlaceholderData");
  const context = requireContext("usePlaceholderData");
  const random = seededRandom(instance.path);

  const pick = <TValue>(values: readonly TValue[]): TValue =>
    values[Math.floor(random() * values.length)];

  return {
    pick,
    name: () => `${pick(firstNames)} ${pick(lastNames)}`,
    city: () => pick(cities),
    date: (offsetDays = 0) => {
      const base = new Date(Date.UTC(2024, 0, 15));
      base.setUTCDate(base.getUTCDate() + offsetDays);
      return new Intl.DateTimeFormat(context.locale, { dateStyle: "long", timeZone: "UTC" })
        .format(base);
    },
    currency: (amount) =>
      new Intl.NumberFormat(context.locale, { style: "currency", currency: "GBP" })
        .format(amount ?? Math.round(random() * 100_000) / 100),
    sentence: (words = 12) =>
      capitalize(Array.from({ length: words }, () => pick(lorem)).join(" ")) + ".",
    paragraph: (sentences = 3) =>
      Array.from({ length: sentences }, () =>
        capitalize(
          Array.from({ length: 8 + Math.floor(random() * 8) }, () => pick(lorem)).join(" "),
        ) + ".").join(" "),
  };
}

/** The formatting {@linkcode useFormat} hands out, bound to one locale. */
export interface Formatters {
  /**
   * Formats an amount of money.
   *
   * @param amount The amount.
   * @param currency An ISO 4217 code. Defaults to `GBP`.
   * @returns The formatted amount.
   */
  currency(amount: number, currency?: string): string;
  /**
   * Formats a number.
   *
   * @param value The number.
   * @param options Passed straight to `Intl.NumberFormat`.
   * @returns The formatted number.
   */
  number(value: number, options?: Intl.NumberFormatOptions): string;
  /**
   * Formats a date.
   *
   * @param value A date, a parseable string, or a timestamp.
   * @param options Passed straight to `Intl.DateTimeFormat`. Defaults to a long date.
   * @returns The formatted date.
   */
  date(value: Date | string | number, options?: Intl.DateTimeFormatOptions): string;
  /**
   * Joins values the way the locale writes a list.
   *
   * @param values The values to join.
   * @param type Whether the list reads as "and" or "or". Defaults to `conjunction`.
   * @returns The joined list.
   */
  list(values: readonly string[], type?: "conjunction" | "disjunction"): string;
  /**
   * Picks the singular or plural form for a count.
   *
   * @param count How many.
   * @param singular The singular form.
   * @param plural The plural form. Defaults to the singular with an `s`.
   * @returns The form that fits.
   */
  plural(count: number, singular: string, plural?: string): string;
}

/** Locale-aware formatting, so a component does not hand-roll it. */
export function useFormat(locale?: string): Formatters {
  const context = requireContext("useFormat");
  const resolved = locale ?? context.locale;

  return {
    currency: (amount, currency = "GBP") =>
      new Intl.NumberFormat(resolved, { style: "currency", currency }).format(amount),
    number: (value, options) => new Intl.NumberFormat(resolved, options).format(value),
    date: (value, options = { dateStyle: "long" }) =>
      new Intl.DateTimeFormat(resolved, options).format(
        value instanceof Date ? value : new Date(value),
      ),
    list: (values, type = "conjunction") =>
      new Intl.ListFormat(resolved, { style: "long", type }).format(values),
    plural: (count, singular, plural) =>
      count === 1 ? singular : plural ?? `${singular}s`,
  };
}

/**
 * Runs a deriver and hands back what it produced.
 *
 * A deriver is the part of a document that runs where the data is, so this is
 * the one hook whose answer depends on which build is asking. With data in
 * hand it runs and returns the value. Publishing cannot run it — the inputs
 * belong to a request nobody has made — so the invocation is recorded on the
 * node instead and what comes back is a stand-in that knows where the real
 * value will be. Either way the component reads one value and never learns
 * which build it is in.
 *
 * Nobody writes an output key. The result lands under the deriver's own name,
 * numbered if one component asks for the same deriver twice, so the token in
 * a published document writes itself the way a loop's does.
 *
 * @typeParam TInputs What the deriver is called with.
 * @typeParam TResult What it produces.
 * @param deriver The deriver to run, imported from its own file.
 * @param inputs What to call it with.
 * @returns What it produced, or a stand-in for what it will produce. Awaited
 * either way — a deriver may be async, and awaiting a stand-in hands back the
 * stand-in, so one `await` reads correctly in both builds.
 *
 * @example
 * ```tsx
 * import invoiceTotals from "../../derivers/invoice-totals.ts";
 *
 * export const Total: Paragraph = async () => {
 *   const [lines] = useState((data: InvoiceData) => data.lines);
 *   const totals = await useDeriver(invoiceTotals, [lines]);
 *
 *   return <Paragraph>Total due {totals.due}.</Paragraph>;
 * };
 * ```
 */
export function useDeriver<TInputs extends readonly unknown[], TResult>(
  deriver: DeriverModule<TInputs, TResult>,
  inputs: TInputs,
): Promise<TResult> {
  const instance = requireInstance("useDeriver");
  const context = requireContext("useDeriver");
  const output = outputKeyFor(instance, deriver.name);

  // Publishing cannot run it: the inputs belong to a request nobody has made.
  // So the invocation travels on the node instead, and what comes back is a
  // stand-in that already knows where the real value will be — which is what
  // lets `{{derived.…}}` write itself the way a loop's tokens do.
  // Recorded in every mode, because the count is what keeps a second call to
  // the same deriver from landing on the first one's key. Only a publish build
  // carries them onto the node — elsewhere the answer is already in the
  // document and there is nothing for an engine to run.
  instance.derivers.push({
    name: deriver.name,
    output,
    inputs: context.deriverMode === "preserve" ? inputs.map((input) => expr(input)) : [],
  });

  if (context.deriverMode === "preserve") {
    context.derivers.register(deriver.name, deriver.run, deriver.placeholder);

    return Promise.resolve(createDerivedStandIn([output]) as TResult);
  }

  // A preview stands in for anything that costs time, the same way it stands in
  // for a generated node. Waiting here is a person waiting.
  if (context.deriverMode === "placeholder" && deriver.placeholder !== undefined) {
    setPath(context.state.derived, output, deriver.placeholder);

    return Promise.resolve(deriver.placeholder);
  }

  // With data in hand the values are right there, so the deriver is called with
  // them rather than with expressions standing for them. Nothing is recorded on
  // the node: the answer is already in the document.
  return Promise.resolve(deriver.run(inputs as unknown as unknown[], context.state)).then((value) => {
    setPath(context.state.derived, output, value);

    return value as TResult;
  });
}

/**
 * Where one component's call to a deriver puts its result.
 *
 * The deriver's own name, so a published token reads as what produced it. A
 * second call to the same deriver from the same component is numbered, because
 * two results cannot share a key — and numbering by call order keeps the key
 * the same on every build of the same component.
 */
function outputKeyFor(instance: ComponentInstance, name: string): string {
  const already = instance.derivers.filter((invocation) => invocation.name === name).length;

  return already === 0 ? name : `${name}-${already + 1}`;
}

function assign(target: PromptDraft, source: PromptDraft): void {
  for (const [key, value] of Object.entries(source)) {
    if (value !== undefined) {
      target[key as keyof PromptDraft] = value as string;
    }
  }
}

function seededRandom(seed: string): () => number {
  let hash = 2166136261;

  for (let index = 0; index < seed.length; index += 1) {
    hash ^= seed.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }

  return () => {
    hash += 0x6d2b79f5;
    let value = hash;
    value = Math.imul(value ^ (value >>> 15), value | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return ((value ^ (value >>> 14)) >>> 0) / 4294967296;
  };
}

function capitalize(value: string): string {
  return value.charAt(0).toUpperCase() + value.slice(1);
}

const firstNames = ["Avery", "Rowan", "Imani", "Tomas", "Neela", "Fintan", "Marta", "Osei"];
const lastNames = ["Whitfield", "Okonkwo", "Lindqvist", "Bassey", "Moreau", "Ferreira", "Nolan"];
const cities = ["Leeds", "Cork", "Antwerp", "Porto", "Malmo", "Bristol", "Ghent"];
const lorem = [
  "tenancy", "notice", "balance", "review", "account", "period", "statement",
  "renewal", "property", "schedule", "payment", "reference", "agreement",
];

export type { PromptDraft };
