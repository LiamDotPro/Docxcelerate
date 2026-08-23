/**
 * Which addresses exist, and in which languages.
 *
 * The site publishes a route for every page in every language, because a
 * reader following a German link to an untranslated page should get the English
 * text rather than a 404. That is right for readers and wrong for search
 * engines: /de/docs/nodes/graph/ currently serves English prose under
 * `lang="de"`, and five URLs carrying the same English text are five
 * competing duplicates.
 *
 * So indexing follows translation, not routing. A page is indexable in a
 * language when that language actually has the file; anywhere English is
 * standing in, the page still serves — it is simply not offered to search
 * engines, and not listed as a translation of anything.
 *
 * Both the hreflang block and the sitemap read from here, so they cannot
 * disagree about which pages exist.
 */
import { getCollection } from "astro:content";
import { docLocale, docSlug } from "./docs";
import { COMPONENTS, THEMES } from "./registry";
import { DEFAULT_LOCALE, LOCALES, localizePath, type Locale } from "./i18n";

export interface LocalisedRoute {
  /** Canonical, unprefixed path, with the trailing slash the build emits. */
  canonical: string;
  /** Languages with their own copy of this page. English is always among them. */
  locales: Locale[];
}

/** The homepage is written in every language the UI is, so all of them. */
export const HOME_ROUTE: LocalisedRoute = {
  canonical: "/",
  locales: [...LOCALES],
};

/**
 * Every docs page, with the languages it has been translated into.
 *
 * English decides which pages exist — the same rule the sidebar follows — so a
 * translation with no English original is not a page.
 */
export async function docsRoutes(): Promise<LocalisedRoute[]> {
  const all = await getCollection("docs");
  const present = new Map<string, Set<Locale>>();

  for (const entry of all) {
    const slug = docSlug(entry.id);
    const locales = present.get(slug) ?? new Set<Locale>();
    locales.add(docLocale(entry.id));
    present.set(slug, locales);
  }

  return all
    .filter((entry) => docLocale(entry.id) === DEFAULT_LOCALE)
    .map((entry) => {
      const slug = docSlug(entry.id);
      const locales = present.get(slug) ?? new Set<Locale>([DEFAULT_LOCALE]);

      return {
        canonical: `/docs/${slug}/`,
        // Ordered by LOCALES rather than by insertion, so the hreflang block
        // and the sitemap come out in a stable order build to build.
        locales: LOCALES.filter((locale) => locales.has(locale)),
      };
    })
    .sort((a, b) => a.canonical.localeCompare(b.canonical));
}

/**
 * The theme and component galleries, and a page for every entry in them.
 *
 * Offered in all five languages on the same terms as the node catalog: the
 * chrome is translated and the entries are generated from the toolkit, which
 * makes them English everywhere rather than stale in four places.
 */
export function registryRoutes(): LocalisedRoute[] {
  return [
    { canonical: "/themes/", locales: [...LOCALES] },
    ...THEMES.map((theme) => ({
      canonical: `/themes/${theme.id}/`,
      locales: [...LOCALES],
    })),
    { canonical: "/components/", locales: [...LOCALES] },
    ...COMPONENTS.map((component) => ({
      canonical: `/components/${component.id}/`,
      locales: [...LOCALES],
    })),
  ];
}

/** Everything that belongs in the sitemap, homepage first. */
export async function indexableRoutes(): Promise<LocalisedRoute[]> {
  return [HOME_ROUTE, ...registryRoutes(), ...(await docsRoutes())];
}

/**
 * The languages a docs page may be offered in. Used by the page itself, which
 * knows its slug but not what the other languages have.
 */
export async function localesForDocsSlug(slug: string): Promise<Locale[]> {
  const routes = await docsRoutes();
  const route = routes.find((candidate) => candidate.canonical === `/docs/${slug}/`);

  return route?.locales ?? [DEFAULT_LOCALE];
}

/** Absolute URL for a canonical path in one language. */
export function absoluteUrl(path: string, locale: Locale, site: URL | undefined): string {
  return new URL(localizePath(path, locale), site).href;
}
