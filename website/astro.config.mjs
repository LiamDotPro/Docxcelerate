// @ts-check
import { defineConfig } from "astro/config";
import deno from "@deno/astro-adapter";
import mdx from "@astrojs/mdx";
import tailwindcss from "@tailwindcss/vite";
import { DEFAULT_LOCALE, LOCALES, PREFIXED_LOCALES } from "./src/i18n/config.ts";

/**
 * Redirects that exist in English at the root also have to exist under every
 * language prefix — a German reader following an old link should land on the
 * German page, not be bounced back to English.
 *
 * @param {string} from Canonical path being redirected.
 * @param {string} to Canonical path to land on.
 * @returns {Record<string, string>}
 */
const localised = (from, to) =>
  Object.fromEntries([
    [from, to],
    ...PREFIXED_LOCALES.map((locale) => [`/${locale}${from}`, `/${locale}${to}`]),
  ]);

export default defineConfig({
  site: "https://docxcelerate.com",
  // Deployed to Deno Deploy. Every page is prerendered — output stays "static"
  // — so the adapter is here to name the target runtime, not to turn the site
  // into a server. It splits the build into dist/client (what gets served) and
  // dist/server (the Deno entry point), and any future route that opts out of
  // prerendering has somewhere to run.
  adapter: deno(),
  integrations: [mdx()],
  // English keeps the bare paths it has always had; the other four languages
  // are served under a prefix. Nothing that was linkable before this became a
  // multilingual site has moved. The locale list lives in src/i18n/config.ts
  // so the routing and the strings cannot disagree about which languages exist.
  i18n: {
    defaultLocale: DEFAULT_LOCALE,
    locales: [...LOCALES],
    routing: { prefixDefaultLocale: false },
  },
  redirects: {
    ...localised("/docs", "/docs/start-here/"),
    // The page was called "letters and nodes" until the vocabulary settled on
    // documents. Anything already linking to the old slug still lands.
    ...localised(
      "/docs/essentials/letters-and-nodes/",
      "/docs/essentials/documents-and-nodes/",
    ),
    // Static and dynamic was a page about a choice you never made. A component
    // decides what it is from what it supplies, and the package records the
    // answer, so the concept now sits inside the page about writing one.
    ...localised(
      "/docs/essentials/static-and-dynamic/",
      "/docs/writing-nodes/",
    ),
  },
  vite: {
    // Two copies of Vite are installed: Astro pins 6, @tailwindcss/vite is
    // built against 8. The plugin shape is the same at runtime — the build
    // works — but the two Plugin types are nominally different, so the cast is
    // about the duplicate install rather than about this plugin.
    plugins: [/** @type {any} */ (tailwindcss())],
  },
  markdown: {
    shikiConfig: {
      themes: { light: "github-light", dark: "github-dark-dimmed" },
      wrap: false,
    },
  },
});
