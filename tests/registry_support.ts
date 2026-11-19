import { mkdir, readdir, readFile, rename, writeFile } from "node:fs/promises";
import { dirname, join, relative } from "node:path";
import { pathToFileURL } from "node:url";
import ts from "typescript";
import { registryRoot } from "docxcelerate/registry/install";
import { transformDocumentSource } from "docxcelerate/transform";

/**
 * Loading a registry component into a test, the way a project's build would.
 *
 * `tsc` does not apply the document transform a publish build needs, so the
 * registry files it emits for the tests that import their types are not the
 * ones a project would run. The whole of `registry/` is compiled again here,
 * transform included, over the top of them in `dist-tests/registry/`, with its
 * layout kept so the node's `../derivers/` import points where it should. It
 * is written there rather than beside the sources because
 * `npm run clean` removes `dist-tests/`, and anything left in `registry/` would
 * ship in the npm tarball and stop `deno publish` on an unclean tree.
 *
 * Type the result with `typeof import("../registry/…")` on the same file: what
 * this hands back is that module, compiled, and the suite is rooted at the
 * repository so the import resolves.
 */

const OUT = join(process.cwd(), "dist-tests", "registry");
let compiled: Promise<void> | undefined;

async function compileRegistry(): Promise<void> {
  const root = registryRoot();
  const files = (await readdir(root, { recursive: true }))
    .filter((file) => /\.tsx?$/.test(file));

  for (const file of files) {
    const fileName = join(root, file);
    let code = await readFile(fileName, "utf8");

    if (file.endsWith(".tsx")) code = transformDocumentSource(code, { fileName }).code;

    const js = ts.transpileModule(code, {
      fileName,
      compilerOptions: {
        target: ts.ScriptTarget.ES2022,
        module: ts.ModuleKind.ESNext,
        jsx: ts.JsxEmit.ReactJSX,
        jsxImportSource: "docxcelerate/template",
        verbatimModuleSyntax: true,
        rewriteRelativeImportExtensions: true,
      },
    }).outputText;
    const out = join(OUT, relative(root, fileName).replace(/\.tsx?$/, ".js"));

    // Every test file is its own process and compiles into the same folder, so
    // a plain write lets another process import a file it has just emptied,
    // and every export comes back undefined. A rename is atomic.
    const temp = `${out}.${process.pid}.tmp`;
    await mkdir(dirname(out), { recursive: true });
    await writeFile(temp, js);
    await rename(temp, out);
  }
}

/**
 * Imports one registry file, compiled.
 *
 * @param source The path under `registry/`, as the catalog writes it:
 * `components/invoice-totals/invoice-totals.node.tsx`.
 * @returns The module, typed as the caller says it is.
 */
export async function loadRegistryModule<T>(source: string): Promise<T> {
  compiled ??= compileRegistry();
  await compiled;

  return await import(pathToFileURL(join(OUT, source.replace(/\.tsx?$/, ".js"))).href) as T;
}
