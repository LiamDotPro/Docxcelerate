import type { DocumentStyle } from "../domain/types.ts";

/**
 * What a theme is: a named, complete answer to how a document is set.
 *
 * A theme is data, not code. It resolves to a {@linkcode DocumentStyle} — the
 * same object a project could have written by hand — so a themed document
 * carries its own appearance to whatever renders it, and nothing downstream has
 * to know the theme existed. Choosing one is a starting point rather than a
 * commitment: a project spreads the style and overrides the parts it disagrees
 * with.
 *
 * @module
 */

/** What a theme is for, so a catalog can group them. */
export type ThemeCategory =
  | "Correspondence"
  | "Report"
  | "Legal"
  | "Technical"
  | "Marketing";

/** What {@linkcode defineTheme} takes. */
export interface DefineThemeOptions {
  /** The theme's identifier: the URL slug, and what `dxcl add` is given. */
  id: string;
  /** The theme's name, as it is printed. */
  title: string;
  /** One line, for cards and lists. */
  summary: string;
  /** A paragraph: what the theme is for, and what it is not for. */
  detail: string;
  /** What it is for, so a catalog can group it. */
  category: ThemeCategory;
  /** Free-form labels a reader might filter on. */
  tags: string[];
  /**
   * The style, minus its `preset` — which is the theme's own id, so it is
   * filled in rather than repeated.
   */
  style: Omit<DocumentStyle, "preset">;
}

/**
 * A theme: everything about how a document looks, under a name.
 */
export interface Theme {
  /** The theme's identifier: the URL slug, and what `dxcl add` is given. */
  readonly id: string;
  /** The theme's name, as it is printed. */
  readonly title: string;
  /** One line, for cards and lists. */
  readonly summary: string;
  /** A paragraph: what the theme is for, and what it is not for. */
  readonly detail: string;
  /** What it is for, so a catalog can group it. */
  readonly category: ThemeCategory;
  /** Free-form labels a reader might filter on. */
  readonly tags: readonly string[];
  /**
   * Every font the theme asks for, deduplicated.
   *
   * Word substitutes silently for a font the machine does not have, so a theme
   * that says which ones it wants is a theme you can check before you print a
   * thousand copies of it.
   */
  readonly fonts: readonly string[];
  /** The style itself, ready to hand to a project or a renderer. */
  readonly style: DocumentStyle;
}

/**
 * Declares a theme. This is what a file under `themes/` exports, and what the
 * catalog collects.
 *
 * @param options The theme's identity and its style.
 * @returns The theme, with its `preset` and font list filled in.
 *
 * @example
 * ```ts
 * export const houseTheme = defineTheme({
 *   id: "house",
 *   title: "House",
 *   summary: "Our own paper.",
 *   detail: "Body text in the corporate face, headings a shade heavier.",
 *   category: "Correspondence",
 *   tags: ["serif"],
 *   style: { page, typography, paragraph, title, sectionHeading },
 * });
 * ```
 */
export function defineTheme(options: DefineThemeOptions): Theme {
  return {
    id: options.id,
    title: options.title,
    summary: options.summary,
    detail: options.detail,
    category: options.category,
    tags: options.tags,
    fonts: fontsOf(options.style),
    style: { preset: options.id, ...options.style },
  };
}

/**
 * A theme's style with some of it changed.
 *
 * A shallow spread would lose the rest of a group — overriding one margin by
 * hand drops the other three — so each group is merged one level down, which is
 * as deep as the style nests.
 *
 * @param theme The theme to start from.
 * @param overrides The parts to replace.
 * @returns A style, still naming the theme it came from.
 *
 * @example
 * ```ts
 * export const documentStyle = themeStyle(slateReportTheme, {
 *   page: { margins: { topMm: 18, rightMm: 18, bottomMm: 18, leftMm: 18 } },
 * });
 * ```
 */
export function themeStyle(theme: Theme, overrides: ThemeStyleOverrides = {}): DocumentStyle {
  const base = theme.style;

  return {
    ...base,
    ...overrides,
    page: {
      ...base.page,
      ...overrides.page,
      margins: { ...base.page.margins, ...overrides.page?.margins },
    },
    typography: { ...base.typography, ...overrides.typography },
    paragraph: { ...base.paragraph, ...overrides.paragraph },
    title: { ...base.title, ...overrides.title },
    sectionHeading: { ...base.sectionHeading, ...overrides.sectionHeading },
    palette: base.palette || overrides.palette
      ? { ...(base.palette as Required<DocumentStyle>["palette"]), ...overrides.palette }
      : undefined,
    // Merged by name, so adding one block to a theme keeps the rest of them.
    // A named block is replaced whole: half a badge is not a badge.
    blocks: base.blocks || overrides.blocks
      ? { ...base.blocks, ...overrides.blocks } as Required<DocumentStyle>["blocks"]
      : undefined,
  };
}

/** What {@linkcode themeStyle} accepts: any part of a style, at any depth it has. */
export type ThemeStyleOverrides = {
  [TKey in keyof DocumentStyle]?: TKey extends "page"
    ? Partial<Omit<DocumentStyle["page"], "margins">> & {
      margins?: Partial<DocumentStyle["page"]["margins"]>;
    }
    : DocumentStyle[TKey] extends object | undefined ? Partial<NonNullable<DocumentStyle[TKey]>>
    : DocumentStyle[TKey];
};

/** Every font a style names, in the order a reader meets them, without repeats. */
function fontsOf(style: Omit<DocumentStyle, "preset">): string[] {
  return [...new Set([style.typography.bodyFont, style.typography.headingFont])];
}
