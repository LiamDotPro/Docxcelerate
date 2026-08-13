// @ts-check
import { defineConfig } from "astro/config";
import mdx from "@astrojs/mdx";
import tailwindcss from "@tailwindcss/vite";

export default defineConfig({
  site: "https://docxcelerate.dev",
  integrations: [mdx()],
  redirects: {
    "/docs": "/docs/start-here/",
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
