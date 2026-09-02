/**
 * The stand-in values a preview shows where real data has not arrived yet.
 *
 * Everything here is deterministic. The seed is where a component sits in the
 * tree, so the same node shows the same name and the same figures on every
 * build — a preview that reshuffles itself is one nobody can proofread.
 *
 * It lives apart from the hooks because it is a corpus and a generator, not a
 * hook: `usePlaceholderData` is the two-line hook that seeds it from the
 * instance it was called in.
 *
 * @module
 */

/** What {@linkcode createPlaceholderData} hands back. */
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
 * Builds the stand-in values for one component.
 *
 * @param seed What the sequence is drawn from. The component's path, so a node
 * shows the same values on every build.
 * @param locale How dates and amounts are formatted.
 * @returns The stand-ins.
 */
export function createPlaceholderData(seed: string, locale: string): PlaceholderData {
  const random = seededRandom(seed);

  const pick = <TValue>(values: readonly TValue[]): TValue =>
    values[Math.floor(random() * values.length)];

  return {
    pick,
    name: () => `${pick(firstNames)} ${pick(lastNames)}`,
    city: () => pick(cities),
    date: (offsetDays = 0) => {
      const base = new Date(Date.UTC(2024, 0, 15));
      base.setUTCDate(base.getUTCDate() + offsetDays);
      return new Intl.DateTimeFormat(locale, { dateStyle: "long", timeZone: "UTC" })
        .format(base);
    },
    currency: (amount) =>
      new Intl.NumberFormat(locale, { style: "currency", currency: "GBP" })
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

/**
 * A small deterministic generator, seeded from a string.
 *
 * FNV-1a over the seed, then xorshift. Nothing here needs to be unpredictable
 * — it needs to be the same on every machine that previews the same document.
 */
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
  "tenancy",
  "notice",
  "balance",
  "review",
  "account",
  "period",
  "statement",
  "renewal",
  "property",
  "schedule",
  "payment",
  "reference",
  "agreement",
];
