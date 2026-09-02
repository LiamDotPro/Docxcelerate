import { dirname, join } from "node:path";
import {
  ensureDirectory,
  exists,
  isNotFoundError,
  parentPath,
  readDirectoryNames,
  readTextFile,
  writeTextFile,
} from "../internal/fs.ts";
import { fileURLToPath } from "node:url";
import { findRegistryEntry, type RegistryEntry, registryEntry } from "./mod.ts";

/**
 * Installing a registry entry into a document project.
 *
 * A component is copied in as source and re-exported from the project's
 * `nodes/index.ts`; a theme is written out as the project's
 * `document-style.ts`. Both leave files you own — nothing is symlinked, nothing
 * is fetched at build time, and nothing here runs again after the install.
 *
 * This is what `dxcl add` calls. It is exported so a script that sets up twenty
 * projects can call it directly rather than shelling out twenty times.
 *
 * @module
 */

/** What {@linkcode installRegistryEntry} takes. */
export interface InstallRegistryEntryOptions {
  /** The entry to install, with or without a `theme:` / `component:` prefix. */
  ref: string;
  /** The document project to install into — the directory holding `document.project.ts`. */
  projectDir: string;
  /** Overwrite files that are already there. */
  force?: boolean;
}

/** What an install did. */
export interface InstallRegistryEntryResult {
  /** The entry's id, without its prefix. */
  id: string;
  /** Which drawer it came from. */
  kind: RegistryEntry["kind"];
  /** The entry's title, for printing. */
  title: string;
  /** Every file written, project-relative. */
  files: string[];
  /**
   * What the install could not do for you.
   *
   * A component reads fields that have to exist on the project's data type, and
   * nothing here edits your types — guessing at somebody's `types.ts` is how a
   * scaffold earns its reputation. So the fields it needs come back as lines to
   * print, and adding them is the one manual step.
   */
  followUp: string[];
}

/**
 * Installs one theme or component into a document project.
 *
 * @param options The entry, the project, and whether to overwrite.
 * @returns What was written, and what is left for you to do.
 * @throws If the id is not in the registry, the project directory is not one,
 * or a file exists and `force` was not set.
 */
export async function installRegistryEntry(
  options: InstallRegistryEntryOptions,
): Promise<InstallRegistryEntryResult> {
  const entry = registryEntry(options.ref);
  await assertDocumentProject(options.projectDir);

  return entry.kind === "theme"
    ? await installTheme(entry.id, options)
    : await installComponent(entry.id, options);
}

/**
 * The entries a request pulls in, in install order, with duplicates removed.
 *
 * Requirements come before the thing that required them, so a component that
 * leans on another is installed after it — which matters for the export list in
 * `nodes/index.ts` reading top to bottom the way the document does.
 *
 * @param refs What was asked for.
 * @returns Every id to install, in order.
 * @throws If an id is not in the registry.
 */
export function resolveInstallOrder(refs: readonly string[]): string[] {
  const ordered: string[] = [];
  const seen = new Set<string>();

  const visit = (ref: string, chain: string[]): void => {
    const entry = registryEntry(ref);

    if (seen.has(entry.id)) {
      return;
    }

    if (chain.includes(entry.id)) {
      throw new Error(
        `Registry entries require each other in a loop: ${[...chain, entry.id].join(" -> ")}.`,
      );
    }

    if (entry.kind === "component") {
      for (const required of entry.requires) {
        visit(required, [...chain, entry.id]);
      }
    }

    seen.add(entry.id);
    ordered.push(entry.id);
  };

  for (const ref of refs) {
    visit(ref, []);
  }

  return ordered;
}

/**
 * Where the package keeps the component sources.
 *
 * The directory sits beside `dist/` in a published package and beside `src/` in
 * a checkout, which is the same two levels up from this module either way.
 *
 * @returns An absolute path to the `registry/` directory.
 */
export function registryRoot(): string {
  const here = dirname(fileURLToPath(import.meta.url));

  return join(dirname(dirname(here)), "registry");
}

/**
 * The document projects under a directory: the directory itself when it is one,
 * and anything one level down under `documents/`.
 *
 * @param root Where to look. Defaults to the working directory.
 * @returns Project directories, in the order they were found.
 */
export async function findDocumentProjects(root = "."): Promise<string[]> {
  const found: string[] = [];

  if (await isDocumentProject(root)) {
    found.push(root);
  }

  for (const candidate of await directoryNames(join(root, "documents"))) {
    const path = join(root, "documents", candidate);

    if (await isDocumentProject(path)) {
      found.push(path);
    }
  }

  return found;
}

/** Whether a directory holds a `document.project.ts`. */
async function isDocumentProject(path: string): Promise<boolean> {
  return await exists(join(path, "document.project.ts"));
}

async function installComponent(
  id: string,
  options: InstallRegistryEntryOptions,
): Promise<InstallRegistryEntryResult> {
  const entry = findRegistryEntry(`component:${id}`);

  if (!entry || entry.kind !== "component") {
    throw new Error(`No component called "${id}".`);
  }

  const root = registryRoot();
  const written: string[] = [];

  for (const file of entry.files) {
    const contents = await readTextFile(join(root, ...file.source.split("/")));
    await writeProjectFile(
      join(options.projectDir, ...file.target.split("/")),
      contents,
      options.force ?? false,
    );
    written.push(file.target);
  }

  const exportPath = join(options.projectDir, "nodes", "index.ts");
  await appendExports(exportPath, entry.exports, entry.files);
  written.push("nodes/index.ts");

  return {
    id: entry.id,
    kind: "component",
    title: entry.title,
    files: written,
    followUp: [
      ...entry.dataFields.map(
        (field) => `Add ${field.path}: ${field.type} to your data type — ${field.summary}`,
      ),
      `Use it: <${entry.exports[0]} /> in document.tsx, imported from ./nodes/index.ts.`,
      ...(entry.themeHint ? [`Drawn against the ${entry.themeHint} theme: dxcl add ${entry.themeHint}`] : []),
    ],
  };
}

async function installTheme(
  id: string,
  options: InstallRegistryEntryOptions,
): Promise<InstallRegistryEntryResult> {
  const entry = findRegistryEntry(`theme:${id}`);

  if (!entry || entry.kind !== "theme") {
    throw new Error(`No theme called "${id}".`);
  }

  const target = join(options.projectDir, "document-style.ts");

  // Every document project already has a document-style.ts — the scaffold
  // writes one — so refusing to overwrite would mean every theme install
  // needed --force, and always passing --force is how people stop reading what
  // it means. Replacing a file nobody has touched is safe; replacing one
  // somebody has written into is not, and that is the case --force is for.
  await writeProjectFile(
    target,
    documentStyleModule(entry.theme.id, entry.theme.title),
    (options.force ?? false) || await isUneditedStyleFile(target),
  );

  return {
    id: entry.id,
    kind: "theme",
    title: entry.theme.title,
    files: ["document-style.ts"],
    followUp: [
      `Fonts this theme asks for: ${entry.theme.fonts.join(", ")}. Word substitutes silently for one you do not have.`,
      "document.project.ts already passes documentStyle through, so the next preview is themed.",
    ],
  };
}

/**
 * The `document-style.ts` a theme install writes.
 *
 * It imports the theme rather than inlining forty numbers, because a style you
 * cannot read is a style you will not edit. `themeStyle` merges one group at a
 * time, so the commented override below changes a margin without dropping the
 * other three.
 */
function documentStyleModule(id: string, title: string): string {
  const variable = `${camelCase(id)}Theme`;

  return `import type { DocumentStyle } from "docxcelerate/document";
import { ${variable}, themeStyle } from "docxcelerate/themes";

/**
 * ${title} — installed by \`dxcl add ${id}\`.
 *
 * This is your copy. Override what you disagree with: \`themeStyle\` merges one
 * group at a time, so changing a single margin keeps the other three.
 */
export const documentStyle: DocumentStyle = themeStyle(${variable}, {
  // page: { margins: { topMm: 20 } },
  // typography: { bodySizePt: 11 },
});
`;
}

/**
 * Re-exports the installed components from `nodes/index.ts`.
 *
 * Appended rather than rewritten: the file is the project's, and an install
 * that reordered somebody's exports would be editing more than it was asked to.
 * An export already present is left alone, so installing twice is not an error.
 */
async function appendExports(
  indexPath: string,
  exports: string[],
  files: readonly { target: string }[],
): Promise<void> {
  const from = files[0]?.target.replace(/^nodes\//, "./") ?? "./index.ts";
  const line = `export { ${exports.join(", ")} } from "${from}";`;
  const current = await exists(indexPath) ? await readTextFile(indexPath) : "";
  const lines = current.split(/\r?\n/).filter((entry) => entry.trim() !== "");

  if (!lines.includes(line)) {
    lines.push(line);
  }

  await ensureDirectory(parentPath(indexPath));
  await writeTextFile(indexPath, `${lines.join("\n")}\n`);
}

/**
 * Whether a `document-style.ts` is still one the toolkit wrote.
 *
 * Two shapes count: the scaffold's default, which spreads
 * `cleanMinimalDocumentStyle`, and a previous theme install, which says so in
 * its own first line. Anything else has been written into by hand, and a theme
 * install should ask before replacing it.
 */
async function isUneditedStyleFile(path: string): Promise<boolean> {
  if (!await exists(path)) {
    return true;
  }

  const contents = await readTextFile(path);

  return contents.includes("cleanMinimalDocumentStyle") ||
    contents.includes("installed by `dxcl add");
}

async function assertDocumentProject(projectDir: string): Promise<void> {
  if (await isDocumentProject(projectDir)) {
    return;
  }

  throw new Error(
    `${projectDir} is not a document project — no document.project.ts in it. ` +
      "Run dxcl document new to make one, or pass --project with the right directory.",
  );
}

async function writeProjectFile(
  path: string,
  contents: string,
  force: boolean,
): Promise<void> {
  if (!force && await exists(path)) {
    throw new Error(`File already exists: ${path}. Pass --force to overwrite it.`);
  }

  await ensureDirectory(parentPath(path));
  await writeTextFile(path, contents);
}

function camelCase(value: string): string {
  return value
    .split(/[^a-zA-Z0-9]+/)
    .filter(Boolean)
    .map((part, index) =>
      index === 0 ? part.toLowerCase() : `${part.charAt(0).toUpperCase()}${part.slice(1)}`
    )
    .join("");
}

/** The names in a directory, treating a missing one as empty rather than an error. */
async function directoryNames(path: string): Promise<string[]> {
  try {
    return await readDirectoryNames(path);
  } catch (error) {
    if (isNotFoundError(error)) {
      return [];
    }

    throw error;
  }
}

