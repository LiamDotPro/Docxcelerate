// @ts-check
import { defineConfig } from "astro/config";
import deno from "@deno/astro-adapter";
import mdx from "@astrojs/mdx";
import tailwindcss from "@tailwindcss/vite";

export default defineConfig({
  site: "https://docxcelerate.com",
  // Deployed to Deno Deploy. Every page is prerendered — output stays "static"
  // — so the adapter is here to name the target runtime, not to turn the site
  // into a server. It splits the build into dist/client (what gets served) and
  // dist/server (the Deno entry point), and any future route that opts out of
  // prerendering has somewhere to run.
  adapter: deno(),
  integrations: [mdx()],
  redirects: {
    "/docs": "/docs/start-here/",
    // The page was called "letters and nodes" until the vocabulary settled on
    // documents. Anything already linking to the old slug still lands.
    "/docs/essentials/letters-and-nodes/": "/docs/essentials/documents-and-nodes/",
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
