/**
 * Wiring the compiler into the build tools that already do the hard parts.
 *
 * Neither Vite nor esbuild hands a plugin a syntax tree — both take source text
 * and give source text back — so the rewriting itself lives in
 * {@linkcode transformDocumentSource}. What they do supply is everything around
 * it: resolution, watching, caching, sourcemaps, and a place to stand in the
 * pipeline. These are the ten lines that join the two.
 *
 * The types here are structural rather than imported. A plugin object is just a
 * shape both tools accept, and describing it that way keeps Vite and esbuild out
 * of this package's dependencies — a document that is only ever rendered needs
 * neither.
 *
 * @module
 */

import { transformDocumentSource } from "./compile.ts";

/** What both plugins take. */
export interface DocxcelerateTransformPluginOptions {
  /**
   * Which files to compile. Defaults to every `.ts` and `.tsx` that is not in
   * `node_modules`.
   */
  include?: (id: string) => boolean;
  /** Where the emitted helpers are imported from. Defaults to the package. */
  runtimeModule?: string;
}

/** The shape Vite accepts as a plugin. */
export interface VitePluginLike {
  /** The plugin's name, as Vite reports it. */
  name: string;
  /** Runs before the tool's own TypeScript handling. */
  enforce: "pre";
  /**
   * Rewrites one module.
   *
   * @param code The module's source.
   * @param id The module's resolved path.
   * @returns The rewritten source, or `null` to leave it alone.
   */
  transform(code: string, id: string): { code: string; map: null } | null;
}

/** The shape esbuild accepts as a plugin. */
export interface EsbuildPluginLike {
  /** The plugin's name, as esbuild reports it. */
  name: string;
  /**
   * Registers the load hook.
   *
   * @param build The registry esbuild passes in.
   */
  setup(build: EsbuildBuildLike): void;
}

/** The part of esbuild's plugin registry these use. */
export interface EsbuildBuildLike {
  /**
   * Registers a handler for files matching a pattern.
   *
   * @param options Which paths the handler is for.
   * @param handler What to do with one.
   */
  onLoad(
    options: { filter: RegExp },
    handler: (
      args: { path: string },
    ) => Promise<{ contents: string; loader: string } | undefined>,
  ): void;
}

/**
 * Compiles document sources as Vite loads them.
 *
 * `enforce: "pre"` is what matters: this has to see the source before Vite's
 * own esbuild pass strips the types, because an `if` is easier to recognise
 * while the tree still says what everything is.
 *
 * @param options Which files to compile, and where the helpers come from.
 * @returns The plugin, ready for a Vite config's `plugins` array.
 *
 * @example
 * ```ts
 * import { docxcelerateTransform } from "docxcelerate/transform";
 *
 * export default defineConfig({ plugins: [docxcelerateTransform()] });
 * ```
 */
export function docxcelerateTransform(
  options: DocxcelerateTransformPluginOptions = {},
): VitePluginLike {
  const include = options.include ?? defaultInclude;

  return {
    name: "docxcelerate-transform",
    enforce: "pre",
    transform(code, id) {
      if (!include(id)) {
        return null;
      }

      const result = transformDocumentSource(code, {
        fileName: id,
        runtimeModule: options.runtimeModule,
      });

      return result.changed ? { code: result.code, map: null } : null;
    },
  };
}

/**
 * Compiles document sources as esbuild loads them.
 *
 * @param options Which files to compile, and where the helpers come from.
 * @returns The plugin, ready for an esbuild `plugins` array.
 *
 * @example
 * ```ts
 * import { docxcelerateEsbuildTransform } from "docxcelerate/transform";
 *
 * await build({ entryPoints: ["document.project.ts"], plugins: [docxcelerateEsbuildTransform()] });
 * ```
 */
export function docxcelerateEsbuildTransform(
  options: DocxcelerateTransformPluginOptions = {},
): EsbuildPluginLike {
  const include = options.include ?? defaultInclude;

  return {
    name: "docxcelerate-transform",
    setup(build) {
      build.onLoad({ filter: /\.[jt]sx?$/ }, async ({ path }) => {
        if (!include(path)) {
          return undefined;
        }

        const { readFile } = await import("node:fs/promises");
        const source = await readFile(path, "utf8");
        const result = transformDocumentSource(source, {
          fileName: path,
          runtimeModule: options.runtimeModule,
        });

        if (!result.changed) {
          return undefined;
        }

        return {
          contents: result.code,
          loader: path.endsWith("x") ? "tsx" : "ts",
        };
      });
    },
  };
}

function defaultInclude(id: string): boolean {
  return /\.[jt]sx?$/.test(id) && !id.includes("node_modules");
}
