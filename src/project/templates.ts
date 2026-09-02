/**
 * Reading the files a scaffolded project starts life as.
 *
 * The templates are real files under `templates/`, not strings inside this
 * package. They were strings once — 1,679 of the 2,347 lines in `scaffold.ts`
 * were TypeScript inside template literals, which meant the compiler never saw
 * them, the formatter never touched them, and the only test that could be
 * written was one asserting a generated file contained a particular substring.
 * A stale dependency pin lived in there for two minor versions.
 *
 * As files they compile: `tsconfig.templates.json` type-checks them against
 * this package the same way a scaffolded workspace type-checks against the
 * published one.
 *
 * @module
 */

import { readFile } from "node:fs/promises";

/**
 * Where the templates sit relative to this module.
 *
 * Two levels up lands on the package root from `src/project/` and from
 * `dist/project/` alike, so the same path works for the JSR source and the npm
 * build.
 */
const templatesRoot = new URL("../../templates/", import.meta.url);

/**
 * What a template's placeholders are filled with, keyed by the name between the
 * underscores.
 */
export type TemplateValues = Readonly<Record<string, string>>;

/**
 * A placeholder: `__UPPER_SNAKE__`.
 *
 * Upper case only, so a template can still contain a lower-case dunder of its
 * own — `preview/main.ts` sets a `__version__` query parameter, and that is not
 * a placeholder.
 */
const placeholder = /__[A-Z0-9_]+__/g;

/**
 * Reads a template and fills its placeholders.
 *
 * Placeholders are written so the template is still valid source: an identifier
 * where an identifier goes, a string where a string goes. That is what lets the
 * templates be type-checked rather than only pattern-matched.
 *
 * @param path The template's path under `templates/`.
 * @param values What to fill its placeholders with.
 * @returns The finished file.
 * @throws If the template carries a placeholder nothing was given for.
 */
export async function readTemplate(
  path: string,
  values: TemplateValues = {},
): Promise<string> {
  const source = await readFile(new URL(path, templatesRoot), "utf8");

  return source.replace(placeholder, (token) => {
    const name = token.slice(2, -2);

    if (!(name in values)) {
      throw new Error(
        `The template ${path} carries ${token}, and nothing was given to fill it with.`,
      );
    }

    return values[name];
  });
}
