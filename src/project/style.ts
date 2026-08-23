import type { DocumentStyle } from "../domain/types.ts";
import { cleanMinimalTheme } from "../theme/themes/clean-minimal.ts";

/**
 * The style a document falls back to when it declares none.
 *
 * @module
 */

/**
 * A plain, readable style: A4 portrait, one-inch margins, and Aptos at 11pt
 * over Cambria headings.
 *
 * This is what a renderer uses when a document carries no style of its own. It
 * is the Clean Minimal theme's style, exported under the name it had before
 * themes existed — see `docxcelerate/themes` for the rest of them.
 */
export const cleanMinimalDocumentStyle: DocumentStyle = cleanMinimalTheme.style;
