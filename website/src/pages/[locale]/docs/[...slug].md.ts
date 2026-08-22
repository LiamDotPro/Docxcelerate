/**
 * The docs as Markdown in every language except English, which sits at the root
 * instead. Mirrors the routing of the HTML pages exactly, untranslated pages
 * included: a language with no copy of a page still answers at its address,
 * with the English text and the note saying so.
 */
import type { APIRoute } from "astro";
import { docsFor, type LocalisedDoc } from "../../../docs";
import { docMarkdown, MARKDOWN_HEADERS } from "../../../markdown";
import { PREFIXED_LOCALES } from "../../../i18n";

export const prerender = true;

export async function getStaticPaths() {
  const routes = await Promise.all(
    PREFIXED_LOCALES.map(async (locale) => {
      const pages = await docsFor(locale);
      return pages.map((page) => ({
        params: { locale, slug: page.slug },
        props: { page },
      }));
    }),
  );

  return routes.flat();
}

export const GET: APIRoute = ({ props, site }) => {
  const { page } = props as { page: LocalisedDoc };

  return new Response(docMarkdown(page, site), { headers: MARKDOWN_HEADERS });
};
