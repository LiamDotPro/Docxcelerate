// @ts-check
import { defineConfig } from "astro/config";
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

/** @param {string} tagName */
const element = (tagName, children = /** @type {any[]} */ ([]), properties = {}) => ({
  type: "element",
  tagName,
  properties,
  children,
});

/**
 * Wraps every markdown table in a `.table-scroll` box. The box is what scrolls
 * on a phone, so the table can stay a real table at full width: a table set to
 * `display: block` to get that scroll stops stretching its rows to fill it.
 *
 * Code inside a table also gets a line break allowed after each `.` and `/`.
 * Without one, a name like `previewOptions.availableTokens` cannot wrap and
 * takes a phone's whole width; letting it break anywhere instead splits short
 * names mid-word, because the table then sizes that column down to nothing.
 *
 * @returns {(tree: any) => void}
 */
const wrapTables = () => (tree) => {
  /** @param {any} node @param {boolean} inTable */
  const walk = (node, inTable) => {
    node.children?.forEach((/** @type {any} */ child, /** @type {number} */ i) => {
      if (child.type !== "element") return;
      if (child.tagName === "table") {
        node.children[i] = element("div", [child], { className: ["table-scroll"] });
        walk(child, true);
      } else if (inTable && child.tagName === "code") {
        child.children = child.children.flatMap((/** @type {any} */ text) =>
          text.type === "text"
            ? text.value
              .split(/(?<=[^./][./])/)
              .flatMap((/** @type {string} */ part, /** @type {number} */ at) => [
                ...(at > 0 ? [element("wbr")] : []),
                { type: "text", value: part },
              ])
            : [text]
        );
      } else {
        walk(child, inTable);
      }
    });
  };
  walk(tree, false);
};

export default defineConfig({
  site: "https://docxcelerate.com",
  // Served by Cloudflare Workers as static assets (see wrangler.jsonc). Every
  // page is prerendered and there is no adapter, so dist/ is the whole site and
  // nothing runs per request. A route that ever opts out of prerendering would
  // need @astrojs/cloudflare here.
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
    // The two galleries became one page. Both addresses were published, so they
    // land on the section they used to be rather than on a 404.
    ...localised("/themes/", "/registry/#themes"),
    ...localised("/components/", "/registry/#components"),
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
    rehypePlugins: [wrapTables],
    shikiConfig: {
      themes: { light: "github-light", dark: "github-dark-dimmed" },
      wrap: false,
    },
  },
});
