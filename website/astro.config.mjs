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
  },
  vite: {
    plugins: [tailwindcss()],
  },
  markdown: {
    shikiConfig: {
      themes: { light: "github-light", dark: "github-dark-dimmed" },
      wrap: false,
    },
  },
});
