/**
 * llms-full.txt — every English docs page, inlined, in sidebar order.
 *
 * For the reader that would otherwise fetch sixteen URLs. /llms.txt is the
 * index; this is the book. Same source, same transform — see src/markdown.ts.
 */
import type { APIRoute } from "astro";
import { byOrder, docsFor } from "../docs";
import { absolute, docMarkdown, MARKDOWN_HEADERS } from "../markdown";
import { DEFAULT_LOCALE, ui } from "../i18n";
import { INSTALL_COMMAND, SITE } from "../site";

export const prerender = true;

export const GET: APIRoute = async ({ site }) => {
  const t = ui(DEFAULT_LOCALE);
  const pages = (await docsFor(DEFAULT_LOCALE)).sort(byOrder);

  const head = [
    `# ${SITE.name}`,
    "",
    `> ${t.meta.tagline} ${t.meta.description}`,
    "",
    `Install: \`${INSTALL_COMMAND}\`. Needs Node.js 20 or newer.`,
    "",
    `The complete documentation from ${
      absolute("/", DEFAULT_LOCALE, site)
    }, every page in one file.`,
    `The index, with a link per page, is at ${
      absolute("/llms.txt", DEFAULT_LOCALE, site)
    }.`,
  ].join("\n");

  // A rule between pages, because the headings inside them are the page's own
  // and a reader scrolling past one needs to see where the next begins.
  const body = [head, ...pages.map((page) => docMarkdown(page, site))].join("\n\n---\n\n");

  return new Response(`${body}\n`, { headers: MARKDOWN_HEADERS });
};
