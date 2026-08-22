/**
 * Every docs page, again, as plain Markdown.
 *
 * The same page an agent reads as HTML costs it many times the bytes of the
 * Markdown it was written in — nav, sidebar, table of contents, styles and
 * script, none of which is the documentation. So each page is also served at
 * its own address with `.md` on the end, and llms.txt points at those.
 *
 * This is not a second copy of the docs. It is the MDX source with the two
 * things an agent cannot use taken out — the `import` lines, which name files
 * it has no access to, and the Astro components, which are replaced by what
 * they render. Nothing here writes prose; if a page says something, it says it
 * in src/content/docs.
 *
 * The node components expand from the generated catalog, the same source the
 * rendered page reads, so the Markdown carries the real option tables and the
 * real variant sources rather than a hole where a component used to be.
 */
import type { LocalisedDoc } from "./docs";
import { byCategory, nodeType, nodeTypeHref, type NodeType } from "./node-catalog";
import { localizePath, ui, type Locale } from "./i18n";

/**
 * Variant sources, read off disk exactly as NodeVariants.astro reads them —
 * the file shown is the file that ran through the framework to build the
 * preview beside it on the HTML page.
 */
const SOURCES = import.meta.glob("./nodes/**/*.node.tsx", {
  query: "?raw",
  import: "default",
  eager: true,
}) as Record<string, string>;

/** Opens or closes a fenced code block. Everything inside one is left alone. */
const FENCE = /^\s*(?:```|~~~)/;
/** An MDX import — the line, or the first line of a wrapped one. */
const IMPORT = /^import\b/;
/** The end of an import, wrapped or not. */
const IMPORT_END = /\bfrom\s+["'][^"']+["'];?\s*$/;
/** A self-closing component alone on a line, which is how the docs use them. */
const COMPONENT = /^<([A-Z][A-Za-z0-9]*)([^>]*?)\/>\s*$/;

/** One docs page as Markdown, headed by what the HTML page puts in `<head>`. */
export function docMarkdown(page: LocalisedDoc, site: URL | undefined): string {
  const { entry, slug, locale } = page;
  const t = ui(locale);

  const head = [
    `# ${entry.data.title}`,
    "",
    `> ${entry.data.description}`,
    "",
    `Source: ${absolute(`/docs/${slug}/`, locale, site)}`,
  ];

  if (!page.translated) head.push("", t.docs.untranslatedNote);

  return `${head.join("\n")}\n\n${body(page, site)}\n`;
}

/**
 * The MDX body, minus its imports, with its components expanded.
 *
 * A capitalised tag this does not know how to expand throws rather than
 * reaching the output as raw JSX: a page whose Markdown twin silently drops a
 * block is worse than a build that stops and names the component that is new.
 */
function body(page: LocalisedDoc, site: URL | undefined): string {
  const source = page.entry.body;

  if (!source) {
    throw new Error(
      `No body for ${page.entry.id}, so its Markdown twin would be an empty page.`,
    );
  }

  const out: string[] = [];
  let fenced = false;
  let importing = false;

  for (const line of source.split("\n")) {
    if (FENCE.test(line)) {
      fenced = !fenced;
      out.push(line);
      continue;
    }

    if (fenced) {
      out.push(line);
      continue;
    }

    if (importing) {
      importing = !IMPORT_END.test(line);
      continue;
    }

    if (IMPORT.test(line)) {
      importing = !IMPORT_END.test(line);
      continue;
    }

    const component = line.match(COMPONENT);

    if (component) {
      out.push(expand(component[1], component[2], page.locale, site));
      continue;
    }

    out.push(line);
  }

  // Taking the imports out leaves behind the blank lines that separated them.
  return out.join("\n").replace(/\n{3,}/g, "\n\n").trim();
}

/** What a component renders, in Markdown. */
function expand(
  name: string,
  attributes: string,
  locale: Locale,
  site: URL | undefined,
): string {
  const type = attributes.match(/type=["']([^"']+)["']/)?.[1];

  switch (name) {
    case "NodeSummary":
      return nodeSummary(nodeType(named(type, name)), locale);
    case "NodeVariants":
      return nodeVariants(nodeType(named(type, name)), locale);
    case "NodeTypes":
      return nodeCatalogue(locale, site);
    default:
      throw new Error(
        `<${name}> has no Markdown form, so /docs/*.md would drop it. ` +
          `Teach src/markdown.ts what it renders.`,
      );
  }
}

function named(type: string | undefined, component: string): string {
  if (!type) throw new Error(`<${component}> without a type= to expand.`);
  return type;
}

/** The header block on a node type's page: what it is, and every option it takes. */
function nodeSummary(type: NodeType, locale: Locale): string {
  const t = ui(locale);
  const status = type.status === "stable" ? "" : ` _(${t.nodes.status[type.status]})_`;

  const blocks = [`${type.detail}${status}`];

  if (type.renderNote) {
    blocks.push(`> **${t.nodes.renderNote}:** ${type.renderNote}`);
  }

  blocks.push(
    [
      `- **${t.nodes.helpers}:** ${
        type.helpers.length > 0 ? type.helpers.map(code).join(", ") : t.nodes.noHelpers
      }`,
      `- **${t.nodes.kind}:** ${code(type.kind)}`,
      `- **${t.nodes.category}:** ${t.nodes.categories[type.category]}`,
      `- **${t.nodes.resolves}:** ${type.resolves}`,
      `- **${t.nodes.children}:** ${type.children}`,
    ].join("\n"),
  );

  if (type.options.length > 0) {
    blocks.push(
      table(
        [t.nodes.option, t.nodes.type, t.nodes.whatItDoes],
        type.options.map((option) => [
          `${code(option.name)}${option.required ? ` _${t.nodes.required}_` : ""}`,
          code(option.type),
          option.summary,
        ]),
      ),
    );
  }

  return blocks.join("\n\n");
}

/**
 * Every variant of one node type. The HTML page shows a rendered preview beside
 * each; here the resolved JSON stands in for it, which is the more useful half
 * for a reader that is about to write one of these.
 */
function nodeVariants(type: NodeType, locale: Locale): string {
  const t = ui(locale);

  return type.variants
    .map((variant) => {
      const blocks = [
        `### ${variant.title}`,
        variant.summary,
        `Source: ${code(`src/nodes/${variant.sourceFile}`)}`,
        fence("tsx", sourceFor(variant.sourceFile)),
        `**${t.nodes.resolvesTo}**`,
        fence("json", JSON.stringify(variant.resolved, null, 2)),
      ];

      if (variant.prompts && variant.prompts.length > 0) {
        blocks.push(
          `**${t.nodes.endpointAsked}**`,
          variant.prompts
            .map((prompt) => `- ${code(prompt.kind)} — ${prompt.text}`)
            .join("\n"),
        );
      }

      return blocks.join("\n\n");
    })
    .join("\n\n");
}

/** The whole catalog, grouped by category, as the overview pages show it. */
function nodeCatalogue(locale: Locale, site: URL | undefined): string {
  const t = ui(locale);

  return byCategory()
    .map(({ category, types }) => {
      const rows = types.map((type) => {
        const href = nodeTypeHref(type);
        const title = href
          ? `[${type.title}](${absolute(href, locale, site)})`
          : type.title;
        const status =
          type.status === "stable" ? "" : ` _(${t.nodes.status[type.status]})_`;

        return `- **${title}** (${code(type.kind)})${status} — ${type.summary}`;
      });

      return `### ${t.nodes.categories[category]}\n\n${rows.join("\n")}`;
    })
    .join("\n\n");
}

function sourceFor(sourceFile: string): string {
  const source = SOURCES[`./nodes/${sourceFile}`];

  if (source === undefined) {
    throw new Error(`No source at src/nodes/${sourceFile} for a catalogued variant.`);
  }

  return source.trimEnd();
}

function code(value: string): string {
  return "`" + value + "`";
}

function fence(lang: string, source: string): string {
  return "```" + lang + "\n" + source + "\n```";
}

/** A cell carrying a pipe would end its column early, so it is escaped. */
function table(headers: string[], rows: string[][]): string {
  const line = (cells: string[]) =>
    `| ${cells.map((cell) => cell.replaceAll("|", "\\|")).join(" | ")} |`;

  return [line(headers), line(headers.map(() => "---")), ...rows.map(line)].join("\n");
}

/**
 * Absolute, because these files are read away from the site that served them —
 * a relative link in a page an agent has pulled into its context points at
 * nothing.
 */
export function absolute(path: string, locale: Locale, site: URL | undefined): string {
  return new URL(localizePath(path, locale), site).href;
}

/** The `.md` twin of a canonical docs path: /docs/nodes/graph/ → /docs/nodes/graph.md */
export function markdownPath(slug: string): string {
  return `/docs/${slug}.md`;
}

/**
 * The routes are prerendered, so what actually reaches a reader is a `.md` file
 * on disk and whatever content type the static server infers from that. This
 * says it anyway: it is what a dev server sends, and it is what these routes
 * would send if one of them ever stopped being prerendered.
 */
export const MARKDOWN_HEADERS = {
  "content-type": "text/markdown; charset=utf-8",
} as const;
