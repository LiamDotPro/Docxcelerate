import { mkdir, readdir, readFile, writeFile } from "node:fs/promises";
import { dirname, join, relative } from "node:path";
import { pathToFileURL } from "node:url";
import ts from "typescript";
import { registryRoot } from "docxcelerate/registry/install";
import { transformDocumentSource } from "docxcelerate/transform";

/**
 * Loading a registry component into a test, the way a project's build would.
 *
 * The registry sits outside this suite's `rootDir`, and `tsc` would not apply
 * the document transform a publish build needs anyway. So the whole of
 * `registry/` is compiled once, transform included, into `dist-tests/registry/`
 * with its layout kept, which leaves the node's `../derivers/` import pointing
 * where it should. It is written there rather than beside the sources because
 * `npm run clean` removes `dist-tests/`, and anything left in `registry/` would
 * ship in the npm tarball and stop `deno publish` on an unclean tree.
 *
 * The types come from importing the source with `import type`, which `npm run
 * typecheck` checks: what this hands back is the same module, compiled.
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

    await mkdir(dirname(out), { recursive: true });
    await writeFile(out, js);
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
