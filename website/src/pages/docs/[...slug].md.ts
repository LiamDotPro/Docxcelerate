/**
 * The English docs as Markdown — /docs/start-here.md beside /docs/start-here/.
 *
 * Same slug, same content, one dot different. An agent that has the HTML
 * address can reach the Markdown by rule rather than by looking anything up,
 * which is the whole reason the convention is worth following.
 */
import type { APIRoute } from "astro";
import { docsFor, type LocalisedDoc } from "../../docs";
import { docMarkdown, MARKDOWN_HEADERS } from "../../markdown";
import { DEFAULT_LOCALE } from "../../i18n";

export const prerender = true;

export async function getStaticPaths() {
  const pages = await docsFor(DEFAULT_LOCALE);
  return pages.map((page) => ({ params: { slug: page.slug }, props: { page } }));
}

export const GET: APIRoute = ({ props, site }) => {
  const { page } = props as { page: LocalisedDoc };

  return new Response(docMarkdown(page, site), { headers: MARKDOWN_HEADERS });
};
