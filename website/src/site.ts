/**
 * Single source of truth for anything that points outside the marketing site.
 * Swap a URL here and every link follows.
 */

/**
 * The name is the same in every language, so it lives here. The tagline and
 * meta description are not — they are prose, and prose is translated: see
 * `meta` in src/i18n/ui/.
 */
export const SITE = {
  name: "Docxcelerate",
} as const;

/**
 * The headline install command.
 *
 * Note this is `npx docxcelerate`, not `npx dxcl` — npx resolves the package
 * name, and the binary is called dxcl. It works because the package declares
 * exactly one bin, which npx runs regardless of its name. Adding a second bin
 * would break this command; use `npx --package=docxcelerate dxcl` if that ever
 * happens.
 */
export const INSTALL_COMMAND = "npx docxcelerate init my-documents";

/**
 * The managed cloud engine — a hosted generation engine, not a dashboard for
 * the toolkit. It isn't open yet, so this URL currently lands on its sign-in
 * page. Everything the site says about it is marked "soon" until it launches;
 * self-hosting the free engine works today.
 *
 * At launch it will have a free tier requiring only a sign-up — so when
 * CLOUD_AVAILABLE flips to true, the copy around it should lead with that
 * rather than with hosting your own.
 */
export const CLOUD_URL = "https://docxcelerate.thoughtup.deno.net/";
export const CLOUD_AVAILABLE = false;

/**
 * The free, self-hostable generation engine. It is deliberately not part of the
 * npm package — the framework authors and previews documents, the engine runs the
 * prompts, builds the documents and exposes an API for generating more.
 *
 * TODO: set this to the engine's repository or download page. While it is null
 * the site describes the engine but doesn't link to it, rather than linking
 * somewhere wrong.
 */
export const ENGINE_URL: string | null = null;

export const GITHUB_URL = "https://github.com/LiamDotPro/Docxcelerate";
export const NPM_URL = "https://www.npmjs.com/package/docxcelerate";

export interface NavLink {
  /** Key into `nav` in the UI dictionaries — the label is per-language. */
  key: "docs" | "cloud" | "github";
  /**
   * Internal links are written canonically, without a language prefix; the Nav
   * adds the reader's prefix. External URLs are used as they stand.
   */
  href: string;
  external?: boolean;
  /** Renders as a mark rather than a word. */
  icon?: "github";
  /** Shows a "soon" badge beside the label. */
  soon?: boolean;
}

export const NAV: NavLink[] = [
  { key: "docs", href: "/docs/start-here/" },
  { key: "cloud", href: CLOUD_URL, external: true, soon: !CLOUD_AVAILABLE },
  { key: "github", href: GITHUB_URL, external: true, icon: "github" },
];

/**
 * Sidebar group order. Pages declare `group` in frontmatter; unknown groups
 * sort last. A page may also declare a `subgroup`, which nests it one level
 * inside its group — how the per-node-type pages sit under Nodes.
 */
export const DOC_GROUPS = [
  "Start Here",
  "Essentials",
  "Nodes",
  "CLI",
  "Projects",
  "Generation",
  "Reference",
] as const;
