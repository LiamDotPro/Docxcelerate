/**
 * Writing the files a project starts life as.
 *
 * This is what `dxcl init` and `dxcl new` are built on, exported so the same
 * scaffolding can be driven from a script rather than a terminal.
 *
 * @module
 */

import {
  ensureDirectory,
  exists,
  isNotFoundError,
  parentPath,
  readDirectoryNames,
  readTextFile,
  writeTextFile,
} from "../internal/fs.ts";
import { readTemplate, type TemplateValues } from "./templates.ts";
import { version } from "../version.ts";

/** What {@linkcode scaffoldDocumentProject} needs to write a document project. */
export interface ScaffoldDocumentProjectOptions {
  /** The document's name; slugified into the directory name. */
  name: string;
  /** The document's title. Derived from `name` when absent. */
  title?: string;
  /** Where document projects live. Defaults to `documents`. */
  documentsDir?: string;
  /** Overwrite files that are already there, rather than refusing. */
  force?: boolean;
}

/** What {@linkcode scaffoldWorkspaceProject} needs to write a workspace. */
export interface ScaffoldWorkspaceProjectOptions {
  /** The workspace's name; slugified into the directory name. */
  name: string;
  /** Where to create the workspace directory. Defaults to the cwd. */
  parentDir?: string;
  /** The engine the workspace generates against. Defaults to the official one. */
  apiEndpoint?: string;
  /** Whether to write a worked example or an empty workspace. Defaults to `sample`. */
  template?: WorkspaceProjectTemplate;
  /** Overwrite files that are already there, rather than refusing. */
  force?: boolean;
}

/** What {@linkcode generateNodeDefinition} needs to write a node. */
export interface GenerateNodeOptions {
  /** The document project the node belongs to. */
  projectDir: string;
  /** The node's name; slugified into its id and PascalCased into its component. */
  name: string;
  /** Which kind of node to write. Defaults to `paragraph`. */
  type?: GeneratedNodeType;
  /** Overwrite a node file that is already there, rather than refusing. */
  force?: boolean;
}

/** The kinds of node the generator can write. */
export type GeneratedNodeType = "paragraph" | "image" | "graph";

/** Whether a new workspace comes with a worked example or nothing at all. */
export type WorkspaceProjectTemplate = "sample" | "blank";

/** The hosted engine a scaffolded workspace points at by default. */
export const officialDocxcelerateApiServer: string = "https://docxcelerate.thoughtup.deno.net/";
/** The generation endpoint on {@linkcode officialDocxcelerateApiServer}. */
export const officialDocxcelerateApiEndpoint: string =
  `${officialDocxcelerateApiServer}api/letters`;

/** What {@linkcode scaffoldDocumentProject} wrote. */
export interface ScaffoldDocumentProjectResult {
  /** The directory the project was written to. */
  projectDir: string;
  /** The project's entrypoint file. */
  entrypoint: string;
  /** Every file written, in the order they were written. */
  files: string[];
}

/** What {@linkcode scaffoldWorkspaceProject} wrote. */
export interface ScaffoldWorkspaceProjectResult {
  /** The directory the workspace was written to. */
  projectDir: string;
  /** The engine the workspace was pointed at. */
  apiEndpoint: string;
  /** Which template was used. */
  template: WorkspaceProjectTemplate;
  /** Every file written, in the order they were written. */
  files: string[];
}

/** What {@linkcode generateNodeDefinition} wrote. */
export interface GenerateNodeResult {
  /** The node file that was written. */
  filePath: string;
  /** The barrel the node was re-exported from. */
  exportPath: string;
  /** The component's name. */
  componentName: string;
  /** The node's id, as it appears in the document. */
  nodeId: string;
}

/**
 * Writes a document project — types, preview data, style, derivers and a
 * couple of nodes — under the workspace's documents directory.
 *
 * @param options The document's name, and where to put it.
 * @returns Where the project landed, and every file written.
 * @throws If the target exists and `force` was not set.
 */
export async function scaffoldDocumentProject(
  options: ScaffoldDocumentProjectOptions,
): Promise<ScaffoldDocumentProjectResult> {
  const slug = slugify(options.name);
  const title = options.title ?? titleFromSlug(slug);
  const projectDir = joinPath(options.documentsDir ?? "documents", slug);
  const entrypoint = joinPath(projectDir, "document.project.ts");
  const values = { DOCUMENT_ID: slug, DOCUMENT_TITLE: title };
  const files = documentProjectFiles(projectDir, "document", values);

  await ensureScaffoldTarget(projectDir, options.force ?? false);
  await writeScaffoldFiles(files, options.force ?? false);

  return {
    projectDir,
    entrypoint,
    files: files.map((file) => file.path),
  };
}

/**
 * Writes a workspace: the package, the preview app's config, and — unless the
 * `blank` template was asked for — one worked document to read.
 *
 * @param options The workspace's name, and what to put in it.
 * @returns Where the workspace landed, and every file written.
 * @throws If the target exists and `force` was not set.
 */
export async function scaffoldWorkspaceProject(
  options: ScaffoldWorkspaceProjectOptions,
): Promise<ScaffoldWorkspaceProjectResult> {
  const slug = slugify(options.name);
  const projectDir = options.parentDir ? joinPath(options.parentDir, slug) : slug;
  const apiEndpoint = normalizeDocxcelerateApiEndpoint(options.apiEndpoint);
  const template = options.template ?? "sample";
  assertWorkspaceProjectTemplate(template);
  const files = workspaceFiles(projectDir, slug, apiEndpoint, template);

  await ensureScaffoldTarget(projectDir, options.force ?? false);
  await writeScaffoldFiles(files, options.force ?? false);

  return {
    projectDir,
    apiEndpoint,
    template,
    files: files.map((file) => file.path),
  };
}

/**
 * Writes one node into an existing document project and re-exports it from the
 * project's `nodes/index.ts`.
 *
 * @param options The project to write into, and what to call the node.
 * @returns Where the node landed, and the names it was given.
 * @throws If the node file exists and `force` was not set, or the type is not
 * one of {@linkcode GeneratedNodeType}.
 */
export async function generateNodeDefinition(
  options: GenerateNodeOptions,
): Promise<GenerateNodeResult> {
  const nodeId = slugify(options.name);
  const componentName = pascalCase(nodeId);
  const type = options.type ?? "paragraph";
  const nodesDir = joinPath(options.projectDir, "nodes");
  const filePath = joinPath(nodesDir, `${nodeId}.node.tsx`);
  const exportPath = joinPath(nodesDir, "index.ts");

  assertGeneratedNodeType(type);

  await ensureDirectory(nodesDir);
  await writeScaffoldFile(
    filePath,
    await readTemplate(`document/nodes/${type}.node.tsx`, {
      NODE_COMPONENT: componentName,
      NODE_ID: nodeId,
      NODE_TITLE: titleFromSlug(nodeId),
      NODE_TITLE_LOWER: titleFromSlug(nodeId).toLowerCase(),
    }),
    options.force ?? false,
  );
  await appendNodeExport(exportPath, componentName, nodeId);

  return {
    filePath,
    exportPath,
    componentName,
    nodeId,
  };
}

/**
 * A file a scaffold writes: where it goes, and the template it comes from.
 *
 * The path and the template are kept apart from the contents so a scaffold can
 * report what it will write before any of it is read.
 */
interface ScaffoldFile {
  /** Where the file lands, relative to the working directory. */
  path: string;
  /** The template's path under `templates/`. */
  template: string;
  /** What to fill the template's placeholders with. */
  values?: TemplateValues;
}

/**
 * The files a document project is made of.
 *
 * `source` picks which of the two starting points to copy: `document` is the
 * blank one `dxcl document new` writes, and `sample` is the worked example a
 * new workspace comes with. They are separate directories rather than one with
 * switches, because the difference between them is the whole point of each.
 */
function documentProjectFiles(
  projectDir: string,
  source: "document" | "sample",
  values: TemplateValues,
): ScaffoldFile[] {
  const nodes = source === "sample"
    ? ["greeting.node.tsx", "balance-summary.node.tsx"]
    : ["greeting.node.tsx", "intro.node.tsx"];

  return [
    { path: "types.ts", template: "types.ts" },
    { path: "preview-data.ts", template: "preview-data.ts" },
    { path: "document-style.ts", template: "document-style.ts" },
    { path: joinPath("derivers", "index.ts"), template: "derivers/index.ts" },
    ...nodes.map((node) => ({
      path: joinPath("nodes", node),
      template: `nodes/${node}`,
    })),
    { path: joinPath("nodes", "index.ts"), template: "nodes/index.ts" },
    { path: "document.tsx", template: "document.tsx" },
    { path: "document.project.ts", template: "document.project.ts" },
  ].map((file) => ({
    path: joinPath(projectDir, file.path),
    template: `${source}/${file.template}`,
    values,
  }));
}

/** The files a workspace is made of, and the document it starts with. */
function workspaceFiles(
  projectDir: string,
  slug: string,
  apiEndpoint: string,
  template: WorkspaceProjectTemplate,
): ScaffoldFile[] {
  const values: TemplateValues = {
    WORKSPACE_NAME: slug,
    WORKSPACE_TITLE: titleFromSlug(slug),
    API_ENDPOINT: apiEndpoint,
    API_ENDPOINT_LABEL: apiEndpoint || "not configured",
    VERSION: version,
    SAMPLE_NOTE: template === "sample"
      ? "A sample document is available at `documents/welcome/document.project.ts`."
      : "This workspace starts blank. Create a document with `dxcl document new`.",
  };

  // A template whose own name would make a tool act on it carries a .template
  // suffix: a package.json here would be the enclosing package for every file
  // beside it, so `docxcelerate` would stop resolving to this repository; a
  // tsconfig.json would be the one an editor type-checked the templates with;
  // and a .gitignore would apply to templates/ rather than describe one.
  const root = [
    ["package.json", "package.json.template"],
    ["docxcelerate.config.json", "docxcelerate.config.json"],
    ["tsconfig.json", "tsconfig.json.template"],
    ["vite.config.ts", "vite.config.ts"],
    ["index.html", "index.html"],
    [joinPath("preview", "main.ts"), "preview/main.ts"],
    [joinPath("preview", "styles.css"), "preview/styles.css"],
    [".gitignore", "gitignore.template"],
    ["README.md", "README.md"],
  ].map(([path, source]) => ({
    path: joinPath(projectDir, path),
    template: `workspace/${source}`,
    values,
  }));

  if (template === "blank") {
    return [
      ...root,
      { path: joinPath(projectDir, "documents", ".gitkeep"), template: "workspace/gitkeep.template" },
    ];
  }

  return [
    ...root,
    ...documentProjectFiles(joinPath(projectDir, "documents", "welcome"), "sample", {
      ...values,
      DOCUMENT_ID: "welcome",
      DOCUMENT_TITLE: "Welcome",
    }),
  ];
}

/** Reads each template and writes the file it describes. */
async function writeScaffoldFiles(files: readonly ScaffoldFile[], force: boolean): Promise<void> {
  for (const file of files) {
    await writeScaffoldFile(file.path, await readTemplate(file.template, file.values), force);
  }
}

function assertGeneratedNodeType(value: string): asserts value is GeneratedNodeType {
  if (value !== "paragraph" && value !== "image" && value !== "graph") {
    throw new Error(`Unsupported node type: ${value}. Expected "paragraph", "image", or "graph".`);
  }
}

function assertWorkspaceProjectTemplate(
  value: string,
): asserts value is WorkspaceProjectTemplate {
  if (value !== "sample" && value !== "blank") {
    throw new Error(`Unsupported workspace template: ${value}. Expected "sample" or "blank".`);
  }
}

async function ensureScaffoldTarget(projectDir: string, force: boolean): Promise<void> {
  try {
    const entries = await readDirectoryNames(projectDir);

    if (!force && entries.length > 0) {
      throw new Error(`Project directory already exists and is not empty: ${projectDir}`);
    }
  } catch (error) {
    if (isNotFoundError(error)) {
      return;
    }

    throw error;
  }
}

async function writeScaffoldFile(path: string, contents: string, force: boolean): Promise<void> {
  if (!force && await exists(path)) {
    throw new Error(`File already exists: ${path}`);
  }

  await ensureDirectory(parentPath(path));
  await writeTextFile(path, contents);
}

async function appendNodeExport(
  indexPath: string,
  componentName: string,
  nodeId: string,
): Promise<void> {
  const exportLine = `export { ${componentName} } from "./${nodeId}.node.tsx";`;
  const current = await exists(indexPath) ? await readTextFile(indexPath) : "";
  const lines = current.split(/\r?\n/).filter((line) => line.trim() !== "");

  if (!lines.includes(exportLine)) {
    lines.push(exportLine);
  }

  await ensureDirectory(parentPath(indexPath));
  await writeTextFile(indexPath, `${lines.join("\n")}\n`);
}

function slugify(value: string): string {
  const slug = value
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");

  if (slug === "") {
    throw new Error(`Expected a name containing at least one document or number`);
  }

  return slug;
}

/**
 * Turns whatever someone typed for an engine URL into a generation endpoint.
 *
 * A bare server — which is what people paste — gets `/api/letters` appended;
 * a URL that already names a path is left alone. An empty or missing value
 * comes back empty, meaning "use the default".
 *
 * @param value The URL to normalize.
 * @returns The endpoint, or an empty string when there was nothing to
 * normalize or the value was not a URL.
 */
export function normalizeDocxcelerateApiEndpoint(value: string | undefined): string {
  const trimmed = value?.trim() ?? "";
  if (trimmed === "") {
    return "";
  }

  try {
    const url = new URL(trimmed);
    if (url.pathname === "" || url.pathname === "/") {
      url.pathname = "/api/letters";
    }

    return url.toString();
  } catch {
    return trimmed;
  }
}

function pascalCase(value: string): string {
  return value
    .split("-")
    .filter(Boolean)
    .map((part) => `${part.charAt(0).toUpperCase()}${part.slice(1)}`)
    .join("");
}

function titleFromSlug(slug: string): string {
  return slug
    .split("-")
    .filter(Boolean)
    .map((part) => `${part.charAt(0).toUpperCase()}${part.slice(1)}`)
    .join(" ");
}

function joinPath(...parts: string[]): string {
  return parts
    .filter((part) => part !== "")
    .join("/")
    .replaceAll("\\", "/")
    .replace(/\/+/g, "/");
}
