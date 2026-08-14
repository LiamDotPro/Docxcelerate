/**
 * The docs collection, read one language at a time.
 *
 * Layout on disk: English pages sit at the root of src/content/docs, exactly
 * where they always have, and each translation mirrors them under its language
 * code — `essentials/templates.mdx` and `de/essentials/templates.mdx` are the
 * same page in two languages. Nothing had to move to make the site
 * multilingual, so English URLs, git history and edit links are untouched.
 *
 * English is the spine: it decides which pages exist and what order they come
 * in. A language that has not translated a page yet gets the English one,
 * flagged, rather than a gap in the sidebar or a 404 — so a new English page is
 * reachable in all five languages the moment it lands.
 */
import { getCollection, type CollectionEntry } from "astro:content";
import { DEFAULT_LOCALE, localeFromPath, stripLocale, type Locale } from "./i18n";

export type DocEntry = CollectionEntry<"docs">;

/** The language a collection id belongs to. Ids are paths, so this is that. */
export function docLocale(id: string): Locale {
  return localeFromPath(`/${id}`);
}

/**
 * The language-independent slug — what the page is called in every language,
 * and what its URL says after the prefix.
 */
export function docSlug(id: string): string {
  return stripLocale(`/${id}`).replace(/^\//, "");
}

export interface LocalisedDoc {
  /** The entry to render: the translation, or the English original. */
  entry: DocEntry;
  /** Shared across languages, so switching language keeps your place. */
  slug: string;
  locale: Locale;
  /** False when English is standing in for a missing translation. */
  translated: boolean;
}

/** Every docs page, as `locale` should see it. */
export async function docsFor(locale: Locale): Promise<LocalisedDoc[]> {
  const all = await getCollection("docs");

  const english = all.filter((entry) => docLocale(entry.id) === DEFAULT_LOCALE);
  const translations = new Map(
    all
      .filter((entry) => docLocale(entry.id) === locale)
      .map((entry) => [docSlug(entry.id), entry]),
  );

  return english.map((original) => {
    const slug = docSlug(original.id);
    const translation = translations.get(slug);

    return {
      entry: translation ?? original,
      slug,
      locale,
      translated: translation !== undefined,
    };
  });
}

/**
 * Sidebar and route order: `order` from the frontmatter, then title. Kept here
 * so the sidebar and the page list cannot disagree about it.
 */
export function byOrder(a: LocalisedDoc, b: LocalisedDoc): number {
  return (
    a.entry.data.order - b.entry.data.order ||
    a.entry.data.title.localeCompare(b.entry.data.title)
  );
}
