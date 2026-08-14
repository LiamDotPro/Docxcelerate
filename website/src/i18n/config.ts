/**
 * The set of languages the site is published in.
 *
 * English is the default and lives at the root — `/`, `/docs/start-here/`.
 * Every other locale is served under its own prefix — `/de/`, `/de/docs/...`.
 * That keeps every URL that existed before this file did working unchanged.
 *
 * Adding a language is: a code here, an entry in LOCALE_META, a dictionary in
 * ./ui, and a directory of MDX under src/content/docs. Nothing else needs to
 * know how many languages there are.
 */

export const DEFAULT_LOCALE = "en";

export const LOCALES = ["en", "nl", "de", "es", "ru"] as const;

export type Locale = (typeof LOCALES)[number];

export interface LocaleMeta {
  /**
   * What the language calls itself. The switcher lists every language in its
   * own words, so a reader who cannot read the current page can still find
   * their own — the one label that must never be translated.
   */
  native: string;
  /** English name, used in comments, commit messages and build output. */
  english: string;
}

export const LOCALE_META: Record<Locale, LocaleMeta> = {
  en: { native: "English", english: "English" },
  nl: { native: "Nederlands", english: "Dutch" },
  de: { native: "Deutsch", english: "German" },
  es: { native: "Español", english: "Spanish" },
  ru: { native: "Русский", english: "Russian" },
};

/** Locales that carry a URL prefix — everything except the default. */
export const PREFIXED_LOCALES = LOCALES.filter(
  (locale): locale is Exclude<Locale, typeof DEFAULT_LOCALE> => locale !== DEFAULT_LOCALE,
);

export function isLocale(value: unknown): value is Locale {
  return typeof value === "string" && (LOCALES as readonly string[]).includes(value);
}

/**
 * Narrows whatever the caller has — `Astro.currentLocale`, a route param, a
 * value read out of the URL — to a locale, falling back to English rather than
 * throwing. A page that renders in the wrong language is a bug worth seeing;
 * a page that fails to render because a string was `undefined` is worse.
 */
export function asLocale(value: unknown): Locale {
  return isLocale(value) ? value : DEFAULT_LOCALE;
}
