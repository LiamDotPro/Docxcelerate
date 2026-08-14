/**
 * The sitemap, built from the same source as the hreflang block.
 *
 * Hand-written rather than delegated to @astrojs/sitemap, because the routes
 * and the indexable pages are not the same set: every page is routed in five
 * languages, and only the translated ones should be offered. An integration
 * walking the built output would list all 75 docs URLs, 40 of which are English
 * text at a German or Spanish address.
 *
 * Each URL carries the full alternate set for its page, including itself, which
 * is what Google asks for: every language in the group points at every other,
 * or the group is ignored.
 */
import type { APIRoute } from "astro";
import { indexableRoutes } from "../seo";
import { DEFAULT_LOCALE, localizePath } from "../i18n";

export const prerender = true;

export const GET: APIRoute = async ({ site }) => {
  const routes = await indexableRoutes();
  const href = (path: string) => new URL(path, site).href;

  const urls = routes.flatMap((route) =>
    route.locales.map((locale) => {
      const alternates = [
        ...route.locales.map(
          (other) =>
            `    <xhtml:link rel="alternate" hreflang="${other}" href="${
              href(localizePath(route.canonical, other))
            }"/>`,
        ),
        `    <xhtml:link rel="alternate" hreflang="x-default" href="${
          href(localizePath(route.canonical, DEFAULT_LOCALE))
        }"/>`,
      ];

      return [
        "  <url>",
        `    <loc>${href(localizePath(route.canonical, locale))}</loc>`,
        ...alternates,
        "  </url>",
      ].join("\n");
    })
  );

  const body = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9"',
    '        xmlns:xhtml="http://www.w3.org/1999/xhtml">',
    ...urls,
    "</urlset>",
    "",
  ].join("\n");

  return new Response(body, {
    headers: { "content-type": "application/xml; charset=utf-8" },
  });
};
