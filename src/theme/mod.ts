/**
 * The shipped themes, and how to find one.
 *
 * A theme is a complete {@linkcode DocumentStyle} under a name — fonts,
 * colours, page and spacing — so choosing one is a single import rather than
 * forty numbers. Nothing here renders: the theme becomes the style the project
 * declares, and the style travels with the built document.
 *
 * ```ts
 * import { slateReportTheme, themeStyle } from "docxcelerate/themes";
 *
 * export const documentStyle = themeStyle(slateReportTheme, {
 *   page: { margins: { topMm: 18 } },
 * });
 * ```
 *
 * `dxcl add slate-report` writes exactly that file into a document project.
 *
 * @module
 */

import type { Theme } from "./theme.ts";
import { boldBriefTheme } from "./themes/bold-brief.ts";
import { cleanMinimalTheme } from "./themes/clean-minimal.ts";
import { legalSerifTheme } from "./themes/legal-serif.ts";
import { slateReportTheme } from "./themes/slate-report.ts";
import { warmLetterTheme } from "./themes/warm-letter.ts";

export * from "./theme.ts";
export { boldBriefTheme, cleanMinimalTheme, legalSerifTheme, slateReportTheme, warmLetterTheme };

/**
 * The colours a chart draws its series in when a theme names none.
 *
 * Exported because a theme writing its own `palette.series` wants to see what
 * it is replacing, and because extending the shipped set is more often right
 * than starting from nothing: the order is what keeps neighbouring series
 * telling apart for a colourblind reader, and it is not a preference.
 */
export { DEFAULT_SERIES_COLORS } from "../render/chart_part.ts";

/**
 * Every theme the package ships, in the order a catalog lists them: the
 * default first, then by how far each one departs from it.
 */
export const THEMES: readonly Theme[] = [
  cleanMinimalTheme,
  slateReportTheme,
  warmLetterTheme,
  legalSerifTheme,
  boldBriefTheme,
];

/**
 * The id of a shipped theme.
 *
 * A document's `style.preset` is a plain string — a project may set a theme of
 * its own — but anywhere the package itself names one, this is the type, so a
 * typo is a build error rather than a lookup that returns nothing.
 */
export type ShippedThemeId =
  | "clean-minimal"
  | "slate-report"
  | "warm-letter"
  | "legal-serif"
  | "bold-brief";

/** Every shipped theme id, for a CLI listing or a static route. */
export const THEME_IDS: readonly ShippedThemeId[] = THEMES.map(
  (theme) => theme.id as ShippedThemeId,
);

/**
 * Looks a theme up by id.
 *
 * @param id The theme's id, as it appears in a URL or on the command line.
 * @returns The theme, or `undefined` when nothing carries that id.
 */
function findTheme(id: string): Theme | undefined {
  return THEMES.find((theme) => theme.id === id);
}

/**
 * Looks a theme up by id, insisting there is one.
 *
 * @param id The theme's id.
 * @returns The theme.
 * @throws If no shipped theme carries that id — a page or a command naming a
 * theme that left the catalog should fail rather than render empty.
 */
export function themeById(id: string): Theme {
  const theme = findTheme(id);

  if (!theme) {
    throw new Error(
      `Unknown theme "${id}". Shipped themes: ${THEME_IDS.join(", ")}.`,
    );
  }

  return theme;
}

/** The themes in one category, in catalog order. */
export function themesByCategory(category: Theme["category"]): Theme[] {
  return THEMES.filter((theme) => theme.category === category);
}
