/**
 * llms.txt — the site's table of contents, for a reader that arrived without
 * one.
 *
 * An agent asked to use Docxcelerate has two ways to find out what it is: guess
 * from a package name, or read something. This is the something. It names every
 * docs page, in the order the sidebar puts them, and links the Markdown twin of
 * each rather than the HTML, so following a link costs a page of prose instead
 * of a page of markup.
 *
 * English only, deliberately. The file exists to be read by a machine that will
 * then write TypeScript, and the English pages are the ones that decide what
 * exists — see src/docs.ts. /llms-full.txt beside it is the same set inlined.
 *
 * The convention is llmstxt.org: an H1, a blockquote, prose, then H2 sections
 * of links. Nothing about it is standardised beyond that, and no search engine
 * has committed to reading it; coding agents fetching library docs do, which is
 * the audience that matters here.
 */
import type { APIRoute } from "astro";
import { byOrder, docsFor } from "../docs";
import { absolute, markdownPath, MARKDOWN_HEADERS } from "../markdown";
import { DEFAULT_LOCALE, ui } from "../i18n";
import {
  DOC_GROUPS,
  GITHUB_URL,
  INSTALL_COMMAND,
  JSR_URL,
  NPM_URL,
  SITE,
} from "../site";

export const prerender = true;

export const GET: APIRoute = async ({ site }) => {
  const t = ui(DEFAULT_LOCALE);
  const url = (path: string) => absolute(path, DEFAULT_LOCALE, site);
  const pages = (await docsFor(DEFAULT_LOCALE)).sort(byOrder);

  const groups = DOC_GROUPS.map((group) => {
    const entries = pages.filter((page) => page.entry.data.group === group);

    return entries.length === 0 ? null : [
      `## ${t.docs.groups[group]}`,
      "",
      ...entries.map(
        (page) =>
          `- [${page.entry.data.title}](${url(markdownPath(page.slug))}): ${
            page.entry.data.description
          }`,
      ),
    ].join("\n");
  }).filter((group) => group !== null);

  const body = [
    `# ${SITE.name}`,
    "",
    `> ${t.meta.tagline} ${t.meta.description}`,
    "",
    [
      `- Install: \`${INSTALL_COMMAND}\`. Needs Node.js 20 or newer and nothing else.`,
      "- TypeScript and JSX. Documents are components; the output is a real .docx file.",
      "- Authoring, preview and DOCX packing run locally. No account, no service call.",
      `- Every page below is also HTML at the same address without the \`.md\`.`,
      `- The agent skill is at ${url("/skill.md")} — drop it in a project and its`,
      "  conventions travel with the repository.",
    ].join("\n"),
    "",
    ...groups.flatMap((group) => [group, ""]),
    "## Project",
    "",
    [
      `- [Source](${GITHUB_URL})`,
      `- [npm](${NPM_URL})`,
      `- [JSR](${JSR_URL})`,
      `- [Every page inlined](${url("/llms-full.txt")}): the whole documentation as one file.`,
    ].join("\n"),
    "",
  ].join("\n");

  return new Response(body, { headers: MARKDOWN_HEADERS });
};
