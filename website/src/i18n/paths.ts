/**
 * Translating a path between locales.
 *
 * Slugs are deliberately *not* translated: `/de/docs/nodes/graph/` is the
 * German text at the English address. It means every page has the same address
 * in all five languages, so the language switcher is a pure prefix swap, an
 * internal link written once works everywhere, and a link someone shares
 * survives being opened by a reader with a different language. hreflang tells
 * search engines which is which.
 */
import { DEFAULT_LOCALE, LOCALES, isLocale, type Locale } from "./config";

/** Leading slash, no double slashes. Trailing slash is left as given. */
function normalize(path: string): string {
  return `/${path.replace(/^\/+/, "")}`.replace(/\/{2,}/g, "/");
}

/**
 * The locale a pathname is served in — the prefix if it has one, English
 * otherwise. `/de/docs/` is German; `/docs/` and `/deno/` are both English.
 */
export function localeFromPath(pathname: string): Locale {
  const first = normalize(pathname).split("/")[1];
  return isLocale(first) && first !== DEFAULT_LOCALE ? first : DEFAULT_LOCALE;
}

/**
 * The canonical (English, unprefixed) form of a pathname. This is the shape
 * every internal link is written in, and the key pages are matched on when
 * switching language.
 */
export function stripLocale(pathname: string): string {
  const path = normalize(pathname);
  const [, first, ...rest] = path.split("/");

  if (!isLocale(first) || first === DEFAULT_LOCALE) return path;

  const remainder = rest.join("/");
  return remainder === "" ? "/" : normalize(remainder);
}

/**
 * A canonical path, addressed in `locale`. Idempotent on paths that already
 * carry a prefix, so callers can pass `Astro.url.pathname` straight in.
 */
export function localizePath(path: string, locale: Locale): string {
  const canonical = stripLocale(path);
  if (locale === DEFAULT_LOCALE) return canonical;
  return canonical === "/" ? `/${locale}/` : normalize(`/${locale}${canonical}`);
}

/**
 * Every language this path is published in, for the `hreflang` block. English
 * is also emitted as `x-default`: it is what a reader with no matching
 * language should land on.
 */
export function alternates(path: string): { locale: Locale; path: string }[] {
  return LOCALES.map((locale) => ({ locale, path: localizePath(path, locale) }));
}
