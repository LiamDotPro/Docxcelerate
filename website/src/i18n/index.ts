/**
 * The one import site for anything translated.
 *
 * A page or component asks for `ui(locale)` and reads strings off it. Nothing
 * outside this directory needs to know which locales exist or where their
 * strings live.
 */
import { en, type UiStrings } from "./ui/en";
import { nl } from "./ui/nl";
import { de } from "./ui/de";
import { es } from "./ui/es";
import { ru } from "./ui/ru";
import type { Locale } from "./config";

export * from "./config";
export * from "./paths";
export type { UiStrings };

const DICTIONARIES: Record<Locale, UiStrings> = { en, nl, de, es, ru };

/** Every string, in one language. */
export function ui(locale: Locale): UiStrings {
  return DICTIONARIES[locale];
}

/**
 * Substitutes `{name}` placeholders. Deliberately tiny: the site has a handful
 * of interpolated strings — a year, two counts, a hash, a node title — and none
 * of them need pluralisation rules or a formatting library.
 *
 * An unknown placeholder is left as it stands rather than blanked, so a typo
 * shows up as `{yaer}` on the page instead of disappearing silently.
 */
export function fill(template: string, values: Record<string, string | number>): string {
  return template.replace(/\{(\w+)\}/g, (match, key: string) =>
    key in values ? String(values[key]) : match,
  );
}
