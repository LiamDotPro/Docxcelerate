/**
 * The toolkit's own version.
 *
 * It is here rather than read from a config file because neither config file is
 * available everywhere this runs: `package.json` does not ship to JSR, and
 * `deno.json` does not ship to npm. A constant ships with the source to both.
 *
 * `tests/package.test.ts` fails when this and the two config files disagree, so
 * a release cannot go out with three different answers.
 *
 * @module
 */

/** The version of this package, as published. */
export const version = "0.4.4";
