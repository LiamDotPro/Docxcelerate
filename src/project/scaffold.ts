export interface ScaffoldDocumentProjectOptions {
  name: string;
  title?: string;
  documentsDir?: string;
  /**
   * What `documentsDir` was called before the vocabulary settled on documents.
   * Still honoured so an existing script keeps scaffolding where it expects.
   *
   * @deprecated Use {@link ScaffoldDocumentProjectOptions.documentsDir}.
   */
  lettersDir?: string;
  force?: boolean;
}

export interface ScaffoldWorkspaceProjectOptions {
  name: string;
  parentDir?: string;
  apiEndpoint?: string;
  template?: WorkspaceProjectTemplate;
  force?: boolean;
}

export interface GenerateNodeOptions {
  projectDir: string;
  name: string;
  type?: GeneratedNodeType;
  mode?: GeneratedNodeMode;
  kind?: GeneratedNodeMode;
  force?: boolean;
}

export type GeneratedNodeType = "paragraph" | "image" | "graph";

export type GeneratedNodeMode = "static" | "dynamic";

export type WorkspaceProjectTemplate = "sample" | "blank";

export const officialDocxcelerateApiServer: string = "https://docxcelerate.thoughtup.deno.net/";
export const officialDocxcelerateApiEndpoint: string =
  `${officialDocxcelerateApiServer}api/letters`;

export interface ScaffoldDocumentProjectResult {
  projectDir: string;
  entrypoint: string;
  files: string[];
}

export interface ScaffoldWorkspaceProjectResult {
  projectDir: string;
  apiEndpoint: string;
  template: WorkspaceProjectTemplate;
  files: string[];
}

export interface GenerateNodeResult {
  filePath: string;
  exportPath: string;
  componentName: string;
  nodeId: string;
}

export async function scaffoldDocumentProject(
  options: ScaffoldDocumentProjectOptions,
): Promise<ScaffoldDocumentProjectResult> {
  const slug = slugify(options.name);
  const title = options.title ?? titleFromSlug(slug);
  const projectDir = joinPath(options.documentsDir ?? options.lettersDir ?? "documents", slug);
  const entrypoint = joinPath(projectDir, "document.project.ts");
  const files = [
    {
      path: joinPath(projectDir, "types.ts"),
      contents: typesTemplate(),
    },
    {
      path: joinPath(projectDir, "preview-data.ts"),
      contents: previewDataTemplate(),
    },
    {
      path: joinPath(projectDir, "document-style.ts"),
      contents: documentStyleTemplate(),
    },
    {
      path: joinPath(projectDir, "derivers", "index.ts"),
      contents: deriversIndexTemplate(),
    },
    {
      path: joinPath(projectDir, "nodes", "greeting.node.tsx"),
      contents: greetingNodeTemplate(),
    },
    {
      path: joinPath(projectDir, "nodes", "intro.node.tsx"),
      contents: introNodeTemplate(),
    },
    {
      path: joinPath(projectDir, "nodes", "index.ts"),
      contents: nodesIndexTemplate(),
    },
    {
      path: joinPath(projectDir, "document.tsx"),
      contents: documentTemplate({ id: slug, title }),
    },
    {
      path: entrypoint,
      contents: projectTemplate({ id: slug, title }),
    },
  ];

  await ensureScaffoldTarget(projectDir, options.force ?? false);

  for (const file of files) {
    await writeScaffoldFile(file.path, file.contents, options.force ?? false);
  }

  return {
    projectDir,
    entrypoint,
    files: files.map((file) => file.path),
  };
}

export async function scaffoldWorkspaceProject(
  options: ScaffoldWorkspaceProjectOptions,
): Promise<ScaffoldWorkspaceProjectResult> {
  const slug = slugify(options.name);
  const projectDir = options.parentDir ? joinPath(options.parentDir, slug) : slug;
  const apiEndpoint = normalizeDocxcelerateApiEndpoint(options.apiEndpoint);
  const template = options.template ?? "sample";
  assertWorkspaceProjectTemplate(template);
  const files = [
    {
      path: joinPath(projectDir, "package.json"),
      contents: workspacePackageJsonTemplate(slug),
    },
    {
      path: joinPath(projectDir, "docxcelerate.config.json"),
      contents: workspaceConfigTemplate(apiEndpoint),
    },
    {
      path: joinPath(projectDir, "tsconfig.json"),
      contents: workspaceTsconfigTemplate(),
    },
    {
      path: joinPath(projectDir, "vite.config.ts"),
      contents: workspaceViteConfigTemplate(),
    },
    {
      path: joinPath(projectDir, "index.html"),
      contents: workspaceIndexHtmlTemplate(slug),
    },
    {
      path: joinPath(projectDir, "preview", "main.ts"),
      contents: workspacePreviewMainTemplate(),
    },
    {
      path: joinPath(projectDir, "preview", "styles.css"),
      contents: workspacePreviewStylesTemplate(),
    },
    {
      path: joinPath(projectDir, ".gitignore"),
      contents: workspaceGitignoreTemplate(),
    },
    {
      path: joinPath(projectDir, "README.md"),
      contents: workspaceReadmeTemplate(slug, apiEndpoint, template),
    },
    ...workspaceDocumentFiles(projectDir, template),
  ];

  await ensureScaffoldTarget(projectDir, options.force ?? false);

  for (const file of files) {
    await writeScaffoldFile(file.path, file.contents, options.force ?? false);
  }

  return {
    projectDir,
    apiEndpoint,
    template,
    files: files.map((file) => file.path),
  };
}

export async function generateNodeDefinition(
  options: GenerateNodeOptions,
): Promise<GenerateNodeResult> {
  const nodeId = slugify(options.name);
  const componentName = pascalCase(nodeId);
  const type = options.type ?? "paragraph";
  const mode = options.mode ?? options.kind ?? "static";
  const nodesDir = joinPath(options.projectDir, "nodes");
  const filePath = joinPath(nodesDir, `${nodeId}.node.tsx`);
  const exportPath = joinPath(nodesDir, "index.ts");

  assertGeneratedNodeType(type);
  assertGeneratedNodeMode(mode);

  await ensureDirectory(nodesDir);
  await writeScaffoldFile(
    filePath,
    nodeTemplate({ componentName, nodeId, type, mode }),
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

function assertGeneratedNodeType(value: string): asserts value is GeneratedNodeType {
  if (value !== "paragraph" && value !== "image" && value !== "graph") {
    throw new Error(`Unsupported node type: ${value}. Expected "paragraph", "image", or "graph".`);
  }
}

function assertGeneratedNodeMode(value: string): asserts value is GeneratedNodeMode {
  if (value !== "static" && value !== "dynamic") {
    throw new Error(`Unsupported node mode: ${value}. Expected "static" or "dynamic".`);
  }
}

function assertWorkspaceProjectTemplate(
  value: string,
): asserts value is WorkspaceProjectTemplate {
  if (value !== "sample" && value !== "blank") {
    throw new Error(`Unsupported workspace template: ${value}. Expected "sample" or "blank".`);
  }
}

function nodeTemplate(options: {
  componentName: string;
  nodeId: string;
  type: GeneratedNodeType;
  mode: GeneratedNodeMode;
}): string {
  if (options.type === "image") {
    return options.mode === "dynamic"
      ? dynamicImageNodeTemplate(options)
      : staticImageNodeTemplate(options);
  }

  if (options.type === "graph") {
    return options.mode === "dynamic"
      ? dynamicGraphNodeTemplate(options)
      : staticGraphNodeTemplate(options);
  }

  return options.mode === "dynamic"
    ? dynamicParagraphNodeTemplate(options)
    : staticParagraphNodeTemplate(options);
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

async function ensureDirectory(path: string): Promise<void> {
  if (path === "" || path === ".") {
    return;
  }

  await mkdir(path);
}

async function exists(path: string): Promise<boolean> {
  try {
    await stat(path);
    return true;
  } catch (error) {
    if (isNotFoundError(error)) {
      return false;
    }

    throw error;
  }
}

async function readDirectoryNames(path: string): Promise<string[]> {
  const { readdir } = await import("node:fs/promises");
  return await readdir(path);
}

async function readTextFile(path: string): Promise<string> {
  const { readFile } = await import("node:fs/promises");
  return await readFile(path, "utf8");
}

async function writeTextFile(path: string, contents: string): Promise<void> {
  const { writeFile } = await import("node:fs/promises");
  await writeFile(path, contents, "utf8");
}

async function mkdir(path: string): Promise<void> {
  const { mkdir: nodeMkdir } = await import("node:fs/promises");
  await nodeMkdir(path, { recursive: true });
}

async function stat(path: string): Promise<void> {
  const { stat: nodeStat } = await import("node:fs/promises");
  await nodeStat(path);
}

function isNotFoundError(error: unknown): boolean {
  return Boolean(
    error && typeof error === "object" && "code" in error && error.code === "ENOENT",
  );
}

function typesTemplate(): string {
  return `export interface DocumentData {
  recipientName: string;
  city: string;
}
`;
}

function previewDataTemplate(): string {
  return `import type { DocumentData } from "./types.ts";

export const previewData: DocumentData = {
  recipientName: "Avery",
  city: "Berlin",
};
`;
}

function sampleTypesTemplate(): string {
  return `export interface DocumentData {
  recipientName: string;
  city: string;
  balanceDue: number;
}
`;
}

function samplePreviewDataTemplate(): string {
  return `import type { DocumentData } from "./types.ts";

export const previewData: DocumentData = {
  recipientName: "Avery",
  city: "Berlin",
  balanceDue: 128.42,
};
`;
}

function documentStyleTemplate(): string {
  return `import { cleanMinimalDocumentStyle, type DocumentStyle } from "docxcelerate/document";

export const documentStyle: DocumentStyle = {
  ...cleanMinimalDocumentStyle,
  page: {
    ...cleanMinimalDocumentStyle.page,
    margins: {
      topMm: 25.4,
      rightMm: 25.4,
      bottomMm: 25.4,
      leftMm: 25.4,
    },
  },
};
`;
}

function workspacePackageJsonTemplate(name: string): string {
  return `${
    JSON.stringify(
      {
        name,
        private: true,
        type: "module",
        scripts: {
          dev: "vite --host 127.0.0.1 --port 4507",
          "document:new": "dxcl document new",
          "document:node": "dxcl document node",
          "documents:check": "tsc -p tsconfig.json",
        },
        dependencies: {
          docxcelerate: "^0.1.3",
          docx: "^9.6.1",
          "docx-preview": "^0.3.7",
        },
        devDependencies: {
          typescript: "^6.0.3",
          vite: "^8.0.13",
        },
      },
      null,
      2,
    )
  }\n`;
}

function workspaceConfigTemplate(apiEndpoint: string): string {
  return `${
    JSON.stringify(
      {
        schemaVersion: "docxcelerate.config/v0",
        activePreset: "local",
        presets: {
          local: {
            build: {
              outDir: "build",
            },
            upload: {
              endpoint: apiEndpoint,
              method: "POST",
              headers: {},
              body: "stored-document",
            },
          },
        },
      },
      null,
      2,
    )
  }\n`;
}

function workspaceTsconfigTemplate(): string {
  return `{
  "compilerOptions": {
    "target": "ES2022",
    "module": "ESNext",
    "moduleResolution": "Bundler",
    "jsx": "react-jsx",
    "jsxImportSource": "docxcelerate/template",
    "strict": true,
    "noEmit": true,
    "allowImportingTsExtensions": true,
    "skipLibCheck": true,
    "types": ["vite/client"]
  },
  "include": ["documents/**/*.ts", "documents/**/*.tsx", "preview/**/*.ts"]
}
`;
}

function workspaceGitignoreTemplate(): string {
  return `node_modules/
build/
documents/**/build/
dist/
`;
}

function workspaceReadmeTemplate(
  name: string,
  apiEndpoint: string,
  template: WorkspaceProjectTemplate,
): string {
  const sampleText = template === "sample"
    ? `\nA sample document is available at \`documents/welcome/document.project.ts\`.\n`
    : `\nThis workspace starts blank. Create a document with \`dxcl document new\`.\n`;

  return `# ${titleFromSlug(name)}

This is a Docxcelerate document workspace.

Docxcelerate API endpoint: ${apiEndpoint || "not configured"}
${sampleText}

Start the preview:

\`\`\`sh
npm run dev
\`\`\`

Create a document:

\`\`\`sh
dxcl document new
\`\`\`

Generate a node:

\`\`\`sh
dxcl document node
\`\`\`

Add shared runtime values for a document in its \`derivers/index.ts\` file.

Type-check documents:

\`\`\`sh
npm run documents:check
\`\`\`
`;
}

function workspaceDocumentFiles(
  projectDir: string,
  template: WorkspaceProjectTemplate,
): Array<{ path: string; contents: string }> {
  if (template === "blank") {
    return [
      {
        path: joinPath(projectDir, "documents", ".gitkeep"),
        contents: "",
      },
    ];
  }

  const documentDir = joinPath(projectDir, "documents", "welcome");

  return [
    {
      path: joinPath(documentDir, "types.ts"),
      contents: sampleTypesTemplate(),
    },
    {
      path: joinPath(documentDir, "preview-data.ts"),
      contents: samplePreviewDataTemplate(),
    },
    {
      path: joinPath(documentDir, "document-style.ts"),
      contents: documentStyleTemplate(),
    },
    {
      path: joinPath(documentDir, "derivers", "index.ts"),
      contents: sampleDeriversIndexTemplate(),
    },
    {
      path: joinPath(documentDir, "nodes", "greeting.node.tsx"),
      contents: greetingNodeTemplate(),
    },
    {
      path: joinPath(documentDir, "nodes", "balance-summary.node.tsx"),
      contents: sampleBalanceSummaryNodeTemplate(),
    },
    {
      path: joinPath(documentDir, "nodes", "index.ts"),
      contents: sampleNodesIndexTemplate(),
    },
    {
      path: joinPath(documentDir, "document.tsx"),
      contents: sampleDocumentTemplate(),
    },
    {
      path: joinPath(documentDir, "document.project.ts"),
      contents: projectTemplate({ id: "welcome", title: "Welcome" }),
    },
  ];
}

function workspaceIndexHtmlTemplate(name: string): string {
  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>${titleFromSlug(name)} Preview</title>
  </head>
  <body>
    <div id="app"></div>
    <script type="module" src="/preview/main.ts"></script>
  </body>
</html>
`;
}

function workspaceViteConfigTemplate(): string {
  return `import { randomUUID } from "node:crypto";
import { spawn } from "node:child_process";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { defineConfig, type Plugin } from "vite";

interface DocxcelerateConfig {
  schemaVersion: "docxcelerate.config/v0";
  activePreset?: string;
  presets?: Record<string, DocxceleratePreset>;
}

interface DocxceleratePreset {
  build?: {
    outDir?: string;
  };
  upload?: {
    endpoint?: string;
    method?: string;
    headers?: Record<string, string>;
    body?: UploadBodyKind;
  };
}

interface BuildRequestBody {
  artifact?: DocumentProjectArtifactPayload;
  preset?: string;
  upload?: boolean;
}

interface PreviewDocxFile {
  name: string;
  bytes: Uint8Array;
  createdAt: number;
}

/**
 * What the toolkit posts here at build time.
 *
 * Both spellings are optional and either may arrive: an artifact built by an
 * older toolkit carries only the letter names, so every read here falls back
 * from the document name to the letter one.
 */
interface DocumentProjectArtifactPayload {
  manifest: {
    id: string;
    name: string;
    version: string;
    entrypoint?: string;
    previewDocument?: string;
    engineDocument?: string;
    previewLetter?: string;
    engineLetter?: string;
    derivers?: string;
    deriverNames?: string[];
    metadata?: Record<string, unknown>;
  };
  previewDocument?: unknown;
  engineDocument?: unknown;
  previewLetter?: unknown;
  engineLetter?: unknown;
  derivers?: DocumentDeriverBundlePayload;
}

type UploadBodyKind =
  | "document"
  | "letter"
  | "stored-document"
  | "stored-letter"
  | "artifact";

interface DocumentDeriverBundlePayload {
  schemaVersion: "docxcelerate.deriver-bundle/v0";
  format: "esm";
  names: string[];
  source: string;
  entrypoint?: string;
  bundledAt?: string;
}

const previewDocxFiles = new Map<string, PreviewDocxFile>();
const previewDocxTtlMs = 10 * 60 * 1000;
const docxContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

export default defineConfig({
  plugins: [docxcelerateDevPlugin()],
});

function docxcelerateDevPlugin(): Plugin {
  return {
    name: "docxcelerate-dev-api",
    configureServer(server) {
      server.middlewares.use(async (request, response, next) => {
        const url = new URL(request.url ?? "/", "http://127.0.0.1");

        if (!url.pathname.startsWith("/api/docxcelerate/")) {
          next();
          return;
        }

        try {
          const previewFileMatch = url.pathname.match(
            /^\\/api\\/docxcelerate\\/preview-docx\\/([^/]+)\\/[^/]+[.]docx$/,
          );

          if (previewFileMatch && request.method === "GET") {
            const file = previewDocxFiles.get(decodeURIComponent(previewFileMatch[1]));

            if (!file) {
              json(response, 404, { error: "Preview DOCX not found." });
              return;
            }

            docx(response, file);
            return;
          }

          if (url.pathname === "/api/docxcelerate/config" && request.method === "GET") {
            const config = await readConfig();
            json(response, 200, {
              config,
              activePreset: resolvePreset(config),
            });
            return;
          }

          if (url.pathname === "/api/docxcelerate/documents" && request.method === "POST") {
            const body = await readJson(request) as { name?: string; title?: string };
            if (!body.name || body.name.trim() === "") {
              json(response, 400, { error: "Document name is required." });
              return;
            }

            const result = await createDocument(body.name, body.title);

            json(response, 201, {
              ...result,
            });
            return;
          }

          if (url.pathname === "/api/docxcelerate/preview-docx" && request.method === "POST") {
            const bytes = await readBytes(request);

            if (bytes.byteLength === 0) {
              json(response, 400, { error: "DOCX preview body is required." });
              return;
            }

            const id = randomUUID();
            const file: PreviewDocxFile = {
              name: safeDocxFileName(url.searchParams.get("name") ?? "preview.docx"),
              bytes,
              createdAt: Date.now(),
            };
            previewDocxFiles.set(id, file);
            cleanupPreviewDocxFiles();

            json(response, 201, {
              id,
              url: previewDocxUrl(id, file),
            });
            return;
          }

          if (url.pathname === "/api/docxcelerate/build" && request.method === "POST") {
            const body = await readJson(request) as BuildRequestBody;
            if (!body.artifact) {
              json(response, 400, { error: "Artifact is required." });
              return;
            }

            const config = await readConfig();
            const preset = resolvePreset(config, body.preset);
            const result = await writeArtifact(body.artifact, preset);
            const upload = body.upload ? await uploadArtifact(body.artifact, preset) : undefined;

            json(response, 200, {
              ...result,
              upload,
            });
            return;
          }

          json(response, 404, { error: "Not found." });
        } catch (error) {
          json(response, 500, {
            error: error instanceof Error ? error.message : String(error),
          });
        }
      });
    },
  };
}

async function createDocument(
  name: string,
  title: string | undefined,
): Promise<{ projectDir: string; entrypoint: string; importPath: string }> {
  const slug = slugify(name);
  const args = ["document", "new", name, "--dir", "documents"];

  if (title && title.trim() !== "") {
    args.push("--title", title);
  }

  await runDxcl(args);

  const projectDir = "documents/" + slug;
  const entrypoint = projectDir + "/document.project.ts";

  return {
    projectDir,
    entrypoint,
    importPath: "../" + entrypoint,
  };
}

function runDxcl(args: string[]): Promise<void> {
  return new Promise((resolve, reject) => {
    const script = join(process.cwd(), "node_modules", "docxcelerate", "bin", "dxcl.mjs");
    const child = spawn(process.execPath, [script, ...args], {
      cwd: process.cwd(),
    });
    let output = "";

    child.stdout.on("data", (chunk) => {
      output += String(chunk);
    });
    child.stderr.on("data", (chunk) => {
      output += String(chunk);
    });
    child.on("error", reject);
    child.on("exit", (code) => {
      if (code === 0) {
        resolve();
        return;
      }

      reject(new Error(output || "dxcl exited with code " + code));
    });
  });
}

async function readConfig(): Promise<DocxcelerateConfig> {
  const text = await readFile("docxcelerate.config.json", "utf8");
  const config = JSON.parse(text) as DocxcelerateConfig;

  if (config.schemaVersion !== "docxcelerate.config/v0") {
    throw new Error("Unsupported docxcelerate.config.json schemaVersion.");
  }

  return config;
}

function resolvePreset(
  config: DocxcelerateConfig,
  presetName = config.activePreset,
): DocxceleratePreset & { name: string } {
  const presets = config.presets ?? {};
  const name = presetName ?? Object.keys(presets)[0] ?? "local";
  const preset = presets[name];

  if (!preset) {
    throw new Error("Unknown Docxcelerate preset: " + name);
  }

  return {
    name,
    ...preset,
  };
}

async function writeArtifact(
  artifact: DocumentProjectArtifactPayload,
  preset: DocxceleratePreset,
): Promise<{ built: true; outDir: string; files: string[] }> {
  const buildDir = preset.build?.outDir ?? "build";
  const outDir = join(documentDirFromArtifact(artifact), buildDir);
  const manifestPath = join(outDir, "manifest.json");
  const previewPath = join(outDir, artifact.manifest.previewDocument ?? artifact.manifest.previewLetter ?? "preview.json");
  const enginePath = join(outDir, artifact.manifest.engineDocument ?? artifact.manifest.engineLetter ?? "document.json");
  const deriversPath = artifact.derivers
    ? join(outDir, artifact.manifest.derivers ?? "derivers.js")
    : undefined;

  await mkdir(outDir, { recursive: true });
  await writeJson(manifestPath, artifact.manifest);
  await writeJson(previewPath, artifact.previewDocument ?? artifact.previewLetter);
  await writeJson(enginePath, artifact.engineDocument ?? artifact.engineLetter);
  if (deriversPath && artifact.derivers) {
    await writeFile(deriversPath, artifact.derivers.source, "utf8");
  }

  return {
    built: true,
    outDir,
    files: [manifestPath, previewPath, enginePath, deriversPath].filter(
      (path): path is string => Boolean(path),
    ),
  };
}

function documentDirFromArtifact(artifact: DocumentProjectArtifactPayload): string {
  const entrypoint = artifact.manifest.entrypoint?.trim();

  if (entrypoint) {
    return dirname(entrypoint.replace(/^[.][\\\\/]+/, ""));
  }

  return join("documents", slugify(artifact.manifest.id));
}

async function uploadArtifact(
  artifact: DocumentProjectArtifactPayload,
  preset: DocxceleratePreset,
): Promise<{ ok: boolean; status: number; statusText: string; text: string }> {
  const upload = preset.upload;
  const endpoint = upload?.endpoint?.trim();

  if (!endpoint) {
    throw new Error("No upload endpoint configured for this preset.");
  }

  const body = uploadBody(artifact, upload?.body ?? "stored-document");
  const response = await fetch(endpoint, {
    method: upload?.method ?? "POST",
    headers: {
      "content-type": "application/json",
      ...upload?.headers,
    },
    body: JSON.stringify(body),
  });

  return {
    ok: response.ok,
    status: response.status,
    statusText: response.statusText,
    text: await response.text(),
  };
}

function uploadBody(
  artifact: DocumentProjectArtifactPayload,
  bodyType: UploadBodyKind,
): unknown {
  if (bodyType === "artifact") {
    return artifact;
  }

  const engine = artifact.engineDocument ?? artifact.engineLetter;

  if (bodyType === "stored-document" || bodyType === "stored-letter") {
    return {
      id: artifact.manifest.id,
      name: artifact.manifest.name,
      version: artifact.manifest.version,
      // Sent under both names: the engine that receives this is not upgraded
      // at the same moment the toolkit that sends it is.
      document: engine,
      letter: engine,
      derivers: artifact.derivers,
      metadata: artifact.manifest.metadata,
    };
  }

  return engine;
}

function cleanupPreviewDocxFiles(now = Date.now()): void {
  for (const [id, file] of previewDocxFiles) {
    if (now - file.createdAt > previewDocxTtlMs) {
      previewDocxFiles.delete(id);
    }
  }
}

function previewDocxUrl(id: string, file: PreviewDocxFile): string {
  return "/api/docxcelerate/preview-docx/" + encodeURIComponent(id) + "/" +
    encodeURIComponent(file.name);
}

function safeDocxFileName(value: string): string {
  const withoutExtension = value.replace(/[.]docx$/i, "");
  const normalized = withoutExtension.trim().replace(/[^A-Za-z0-9._-]+/g, "-");
  const base = normalized === "" ? "preview" : normalized;

  return base + ".docx";
}

function docx(
  response: import("node:http").ServerResponse,
  file: PreviewDocxFile,
): void {
  response.statusCode = 200;
  response.setHeader("content-type", docxContentType);
  response.setHeader("cache-control", "no-store");
  response.setHeader("content-disposition", 'inline; filename="' + file.name + '"');
  response.setHeader("content-length", String(file.bytes.byteLength));
  response.end(file.bytes);
}

async function writeJson(path: string, value: unknown): Promise<void> {
  await writeFile(path, JSON.stringify(value, null, 2) + "\\n", "utf8");
}

async function readBytes(request: import("node:http").IncomingMessage): Promise<Buffer> {
  const chunks: Uint8Array[] = [];

  for await (const chunk of request) {
    chunks.push(typeof chunk === "string" ? Buffer.from(chunk) : chunk);
  }

  return Buffer.concat(chunks);
}

async function readJson(request: import("node:http").IncomingMessage): Promise<unknown> {
  const text = (await readBytes(request)).toString("utf8");
  return text.trim() === "" ? {} : JSON.parse(text);
}

function json(
  response: import("node:http").ServerResponse,
  statusCode: number,
  value: unknown,
): void {
  response.statusCode = statusCode;
  response.setHeader("content-type", "application/json; charset=utf-8");
  response.end(JSON.stringify(value) + "\\n");
}

function slugify(value: string): string {
  const slug = value
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");

  if (slug === "") {
    throw new Error("Expected a name containing at least one document or number.");
  }

  return slug;
}
`;
}

function workspacePreviewMainTemplate(): string {
  return `import { buildDocument, createDocumentProjectArtifact } from "docxcelerate";
import type { DocumentModel, DocumentProject } from "docxcelerate/document";
import "./styles.css";

interface DocumentProjectModule {
  default?: DocumentProject<unknown>;
  project?: DocumentProject<unknown>;
  documentProject?: DocumentProject<unknown>;
}

interface DocxcelerateConfig {
  schemaVersion: "docxcelerate.config/v0";
  activePreset?: string;
  presets?: Record<string, DocxceleratePreset>;
}

interface DocxceleratePreset {
  name?: string;
  build?: {
    outDir?: string;
  };
  upload?: {
    endpoint?: string;
    method?: string;
    headers?: Record<string, string>;
    body?: UploadBodyKind;
  };
}

type PreviewRenderer = "docx-preview" | "microsoft-office" | "google-docs";

interface PreviewRendererOption {
  value: PreviewRenderer;
  label: string;
}

const projectLoaders = import.meta.glob<DocumentProjectModule>("../documents/**/document.project.ts");
const appElement = document.querySelector<HTMLDivElement>("#app");
const previewRendererStorageKey = "docxcelerate.previewRenderer";
const previewRendererOptions: PreviewRendererOption[] = [
  { value: "docx-preview", label: "docx-preview" },
  { value: "microsoft-office", label: "Microsoft Office" },
  { value: "google-docs", label: "Google Docs" },
];
let route = readRoute();
let config: { config: DocxcelerateConfig; activePreset: DocxceleratePreset } | undefined;
let previewRenderer = readPreviewRenderer();

if (!appElement) {
  throw new Error("Expected #app element");
}

const app = appElement;

void renderApp();

window.addEventListener("hashchange", () => {
  route = readRoute();
  void renderApp();
});

if (import.meta.hot) {
  import.meta.hot.accept(() => {
    window.location.reload();
  });
}

async function renderApp(): Promise<void> {
  const paths = sortedProjectPaths();
  config ??= await loadConfig();

  if (route.view === "home" || paths.length === 0) {
    renderHome(paths);
    return;
  }

  if (!route.path || !projectLoaders[route.path]) {
    navigateHome();
    return;
  }

  renderLoading(paths, route.path);

  try {
    const project = await loadProject(route.path);
    const document = await buildPreviewDocument(project);
    await renderPreview(paths, route.path, project, document);
  } catch (error) {
    renderError(paths, route.path, error);
  }
}

function sortedProjectPaths(): string[] {
  return Object.keys(projectLoaders).sort((left, right) => left.localeCompare(right));
}

async function loadConfig(): Promise<
  { config: DocxcelerateConfig; activePreset: DocxceleratePreset }
> {
  const response = await fetch("/api/docxcelerate/config");
  if (!response.ok) {
    throw new Error("Unable to load docxcelerate.config.json");
  }

  return await response.json();
}

async function loadProject(path: string): Promise<DocumentProject<unknown>> {
  const module = await projectLoaders[path]();
  const project = module.documentProject ?? module.project ?? module.default;

  if (!project) {
    throw new Error("Document project did not export a project: " + path);
  }

  return project;
}

async function buildPreviewDocument(project: DocumentProject<unknown>): Promise<DocumentModel> {
  const document = await buildDocument(project.template, project.previewData, {
    ...project.previewOptions,
    dynamicMode: "placeholder",
  });

  return {
    ...document,
    style: project.style,
    metadata: project.metadata ? { ...project.metadata, ...document.metadata } : document.metadata,
  };
}

function renderHome(paths: string[]): void {
  const shell = createShell({ currentPath: undefined, paths, titleText: "Home" });
  const workspace = shell.querySelector(".workspace");
  const home = document.createElement("section");
  home.className = "home-section";

  const header = document.createElement("div");
  header.className = "home-header";

  const title = document.createElement("h1");
  title.textContent = "Documents";

  const preset = document.createElement("p");
  preset.className = "config-line";
  preset.textContent = configLabel();

  header.append(title, preset);
  home.append(header, renderProjectList(paths), renderNewDocumentForm());

  workspace?.append(home);
  app.replaceChildren(shell);
}

function renderProjectList(paths: string[]): HTMLElement {
  const list = document.createElement("div");
  list.className = "document-list";

  if (paths.length === 0) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "No documents yet";
    list.append(empty);
    return list;
  }

  for (const path of paths) {
    const item = document.createElement("article");
    item.className = "document-card";

    const name = document.createElement("h2");
    name.textContent = labelFromPath(path);

    const entrypoint = document.createElement("p");
    entrypoint.textContent = entrypointFromPath(path);

    const open = document.createElement("button");
    open.type = "button";
    open.className = "button primary-button";
    open.textContent = "Open preview";
    open.addEventListener("click", () => navigatePreview(path));

    item.append(name, entrypoint, open);
    list.append(item);
  }

  return list;
}

function renderNewDocumentForm(): HTMLFormElement {
  const form = document.createElement("form");
  form.className = "new-document-form";

  const title = document.createElement("h2");
  title.textContent = "New document";

  const nameInput = document.createElement("input");
  nameInput.name = "name";
  nameInput.required = true;
  nameInput.placeholder = "case-review";
  nameInput.autocomplete = "off";

  const titleInput = document.createElement("input");
  titleInput.name = "title";
  titleInput.placeholder = "Case Review";
  titleInput.autocomplete = "off";

  const submit = document.createElement("button");
  submit.type = "submit";
  submit.className = "button primary-button";
  submit.textContent = "Add document";

  const message = document.createElement("p");
  message.className = "form-message";

  form.append(title, labeledField("Name", nameInput), labeledField("Title", titleInput), submit, message);
  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    submit.disabled = true;
    message.textContent = "Creating";

    try {
      const body = {
        name: nameInput.value,
        title: titleInput.value || undefined,
      };
      const response = await fetch("/api/docxcelerate/documents", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify(body),
      });
      const result = await response.json();

      if (!response.ok) {
        throw new Error(result.error ?? "Unable to create document.");
      }

      localStorage.setItem("docxcelerate.nextPath", result.importPath);
      window.location.hash = "#preview=" + encodeURIComponent(result.importPath);
      window.location.reload();
    } catch (error) {
      message.textContent = error instanceof Error ? error.message : String(error);
      submit.disabled = false;
    }
  });

  return form;
}

function labeledField(labelText: string, input: HTMLInputElement): HTMLLabelElement {
  const label = document.createElement("label");
  const span = document.createElement("span");
  span.textContent = labelText;
  label.append(span, input);
  return label;
}

function renderLoading(paths: string[], currentPath: string): void {
  const shell = createShell({ currentPath, paths, titleText: "Loading" });
  const loading = document.createElement("div");
  loading.className = "empty-state";
  loading.textContent = "Loading document";
  shell.querySelector(".workspace")?.append(loading);
  app.replaceChildren(shell);
}

function renderError(paths: string[], currentPath: string, error: unknown): void {
  const shell = createShell({ currentPath, paths, titleText: "Preview error" });
  const panel = document.createElement("pre");
  panel.className = "error-panel";
  panel.textContent = error instanceof Error ? error.message : String(error);
  shell.querySelector(".workspace")?.append(panel);
  app.replaceChildren(shell);
}

async function renderPreview(
  paths: string[],
  currentPath: string,
  project: DocumentProject<unknown>,
  document: DocumentModel,
): Promise<void> {
  const shell = createShell({
    currentPath,
    paths,
    titleText: project.name + " " + previewRendererLabel(previewRenderer),
    project,
  });
  shell.querySelector(".workspace")?.append(await renderDocumentPreview(document, previewRenderer));
  app.replaceChildren(shell);
}

function createShell(options: {
  currentPath?: string;
  paths: string[];
  titleText: string;
  project?: DocumentProject<unknown>;
}): HTMLElement {
  const shell = document.createElement("main");
  shell.className = "preview-shell";

  const header = document.createElement("header");
  header.className = "preview-toolbar";

  const brand = document.createElement("button");
  brand.type = "button";
  brand.className = "brand-button";
  brand.textContent = "Docxcelerate";
  brand.addEventListener("click", navigateHome);

  const select = document.createElement("select");
  select.className = "project-select";
  select.setAttribute("aria-label", "Document project");

  for (const path of options.paths) {
    const option = document.createElement("option");
    option.value = path;
    option.textContent = labelFromPath(path);
    option.selected = path === options.currentPath;
    select.append(option);
  }

  select.addEventListener("change", () => {
    navigatePreview(select.value);
  });

  const status = document.createElement("span");
  status.className = "toolbar-status";
  status.textContent = options.titleText;

  const home = document.createElement("button");
  home.type = "button";
  home.className = "button";
  home.textContent = "Home";
  home.addEventListener("click", navigateHome);

  header.append(brand, home, select);

  if (options.project && options.currentPath) {
    header.append(createPreviewRendererSelect());
  }

  header.append(status);

  if (options.project && options.currentPath) {
    header.append(
      actionButton("Build", () =>
        buildProjectFromUi(options.project as DocumentProject<unknown>, options.currentPath as string, false, status)
      ),
      actionButton("Build & upload", () =>
        buildProjectFromUi(options.project as DocumentProject<unknown>, options.currentPath as string, true, status)
      ),
    );
  }

  const workspace = document.createElement("section");
  workspace.className = "workspace";

  shell.append(header, workspace);
  return shell;
}

function createPreviewRendererSelect(): HTMLSelectElement {
  const select = document.createElement("select");
  select.className = "preview-mode-select";
  select.setAttribute("aria-label", "Preview renderer");

  for (const renderer of previewRendererOptions) {
    const option = document.createElement("option");
    option.value = renderer.value;
    option.textContent = renderer.label;
    option.selected = renderer.value === previewRenderer;
    select.append(option);
  }

  select.addEventListener("change", () => {
    previewRenderer = asPreviewRenderer(select.value);
    localStorage.setItem(previewRendererStorageKey, previewRenderer);
    void renderApp();
  });

  return select;
}

function actionButton(label: string, handler: () => Promise<void>): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "button";
  button.textContent = label;

  if (label.includes("upload") && !uploadEndpoint()) {
    button.disabled = true;
    button.title = "Set presets.local.upload.endpoint in docxcelerate.config.json";
  }

  button.addEventListener("click", async () => {
    button.disabled = true;
    try {
      await handler();
    } finally {
      button.disabled = label.includes("upload") && !uploadEndpoint();
    }
  });

  return button;
}

async function buildProjectFromUi(
  project: DocumentProject<unknown>,
  path: string,
  upload: boolean,
  status: HTMLElement,
): Promise<void> {
  status.textContent = upload ? "Building and uploading" : "Building";
  const artifact = await createDocumentProjectArtifact(project, {
    entrypoint: entrypointFromPath(path),
  });
  const response = await fetch("/api/docxcelerate/build", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      artifact,
      upload,
    }),
  });
  const result = await response.json();

  if (!response.ok) {
    throw new Error(result.error ?? "Build failed.");
  }

  if (upload && result.upload && !result.upload.ok) {
    status.textContent = "Upload failed: " + result.upload.status + " " + result.upload.statusText;
    return;
  }

  status.textContent = upload ? "Built and uploaded " + result.outDir : "Built " + result.outDir;
}

async function renderDocumentPreview(
  document: DocumentModel,
  renderer: PreviewRenderer,
): Promise<HTMLElement> {
  const { createDocxBlob } = await import("docxcelerate/docx");
  const documentBlob = await createDocxBlob(document);

  if (renderer === "docx-preview") {
    return await renderClientDocxPreview(document, documentBlob);
  }

  return await renderHostedDocxPreview(document, documentBlob, renderer);
}

async function renderClientDocxPreview(document: DocumentModel, documentBlob: Blob): Promise<HTMLElement> {
  const docxPreview = await import("docx-preview");
  const stage = document.createElement("div");
  stage.className = "docx-preview-stage";

  const frame = document.createElement("iframe");
  frame.className = "docx-preview-frame";
  frame.title = document.title + " DOCX preview";

  const body = document.createElement("body");
  const head = document.createElement("head");

  await docxPreview.renderAsync(documentBlob, body, head, {
    inWrapper: true,
    ignoreLastRenderedPageBreak: false,
  });

  const style = document.createElement("style");
  style.textContent = previewFrameStyles();
  head.append(style);

  const html = document.createElement("html");
  html.append(head, body);

  frame.srcdoc = "<!doctype html>" + html.outerHTML;
  stage.append(frame);

  return stage;
}

async function renderHostedDocxPreview(
  document: DocumentModel,
  documentBlob: Blob,
  renderer: Exclude<PreviewRenderer, "docx-preview">,
): Promise<HTMLElement> {
  const stage = document.createElement("div");
  stage.className = "docx-preview-stage";

  if (isPrivatePreviewHost()) {
    stage.append(renderHostedPreviewNotice(renderer));
    return stage;
  }

  const documentUrl = await uploadPreviewDocx(document, documentBlob);
  const frame = document.createElement("iframe");
  frame.className = "docx-preview-frame";
  frame.title = document.title + " " + previewRendererLabel(renderer);
  frame.src = hostedPreviewUrl(renderer, documentUrl);
  stage.append(frame);

  return stage;
}

function renderHostedPreviewNotice(renderer: PreviewRenderer): HTMLElement {
  const panel = document.createElement("div");
  panel.className = "external-preview-message";

  const title = document.createElement("h2");
  title.textContent = previewRendererLabel(renderer) + " needs a public URL";

  const body = document.createElement("p");
  body.textContent =
    "This preview mode sends a temporary DOCX URL to the external viewer. Run the Vite preview through a public HTTPS tunnel or deployed URL, then reload this mode.";

  panel.append(title, body);
  return panel;
}

async function uploadPreviewDocx(document: DocumentModel, documentBlob: Blob): Promise<string> {
  const response = await fetch(
    "/api/docxcelerate/preview-docx?name=" + encodeURIComponent(document.id + ".docx"),
    {
      method: "POST",
      headers: {
        "content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      },
      body: documentBlob,
    },
  );
  const result = await response.json() as { url?: string; error?: string };

  if (!response.ok || !result.url) {
    throw new Error(result.error ?? "Unable to prepare hosted DOCX preview.");
  }

  return new URL(result.url, window.location.href).toString();
}

function hostedPreviewUrl(
  renderer: Exclude<PreviewRenderer, "docx-preview">,
  documentUrl: string,
): string {
  if (renderer === "microsoft-office") {
    const url = new URL("https://view.officeapps.live.com/op/embed.aspx");
    url.searchParams.set("src", documentUrl);
    return url.toString();
  }

  const url = new URL("https://docs.google.com/gview");
  url.searchParams.set("embedded", "true");
  url.searchParams.set("url", documentUrl);
  url.searchParams.set("__version__", String(Date.now()));
  return url.toString();
}

function isPrivatePreviewHost(hostname = window.location.hostname): boolean {
  const host = hostname.toLowerCase();

  return host === "localhost" ||
    host === "127.0.0.1" ||
    host === "0.0.0.0" ||
    host === "::1" ||
    host.endsWith(".local") ||
    /^10[.]/.test(host) ||
    /^192[.]168[.]/.test(host) ||
    /^172[.](1[6-9]|2[0-9]|3[0-1])[.]/.test(host);
}

function previewFrameStyles(): string {
  return [
    "html, body { margin: 0; min-height: 100%; background: #e6e8ed; }",
    "body { overflow: auto; padding: 32px 0 72px; }",
    ".docx-wrapper { background: transparent !important; padding: 0 !important; }",
    ".docx-wrapper > section.docx { margin: 0 auto 34px !important; box-shadow: 0 1px 2px rgb(15 23 42 / 12%), 0 18px 46px rgb(15 23 42 / 24%); }",
  ].join("\\n");
}

function readPreviewRenderer(): PreviewRenderer {
  return asPreviewRenderer(localStorage.getItem(previewRendererStorageKey));
}

function asPreviewRenderer(value: string | null): PreviewRenderer {
  return value === "microsoft-office" || value === "google-docs" ? value : "docx-preview";
}

function previewRendererLabel(renderer: PreviewRenderer): string {
  return previewRendererOptions.find((option) => option.value === renderer)?.label ?? "docx-preview";
}

function readRoute(): { view: "home" } | { view: "preview"; path: string } {
  const hash = window.location.hash.replace(/^#/, "");
  const nextPath = localStorage.getItem("docxcelerate.nextPath");

  if (nextPath) {
    localStorage.removeItem("docxcelerate.nextPath");
    return {
      view: "preview",
      path: nextPath,
    };
  }

  if (hash.startsWith("preview=")) {
    return {
      view: "preview",
      path: decodeURIComponent(hash.slice("preview=".length)),
    };
  }

  return { view: "home" };
}

function navigateHome(): void {
  window.location.hash = "";
}

function navigatePreview(path: string): void {
  window.location.hash = "#preview=" + encodeURIComponent(path);
}

function labelFromPath(path: string): string {
  return path.replace("../documents/", "").replace("/document.project.ts", "");
}

function entrypointFromPath(path: string): string {
  return path.replace(/^\\.\\.\\//, "");
}

function configLabel(): string {
  const preset = config?.activePreset;
  const name = preset?.name ?? config?.config.activePreset ?? "local";
  const endpoint = uploadEndpoint();

  return endpoint ? "Preset " + name + " -> " + endpoint : "Preset " + name;
}

function uploadEndpoint(): string | undefined {
  return config?.activePreset.upload?.endpoint?.trim() || undefined;
}
`;
}

function workspacePreviewStylesTemplate(): string {
  return `:root {
  color-scheme: light;
  font-family: Aptos, "Segoe UI", Arial, sans-serif;
  background: #e6e8ed;
  color: #1d2430;
  --chrome: #fbfbfd;
  --chrome-border: #d8dce3;
  --workspace: #e6e8ed;
  --page: #ffffff;
  --page-border: #c7ccd6;
  --accent: #2f5fbd;
  --muted: #647083;
}

* {
  box-sizing: border-box;
}

html,
body,
#app {
  margin: 0;
  min-height: 100vh;
}

body {
  overflow: hidden;
  background: var(--workspace);
}

button,
input,
select {
  font: inherit;
}

.preview-shell {
  display: grid;
  grid-template-rows: auto 1fr;
  height: 100vh;
  min-width: 900px;
}

.preview-toolbar {
  display: grid;
  grid-template-columns: auto auto minmax(220px, 360px) minmax(160px, 220px) minmax(0, 1fr) auto auto;
  align-items: center;
  gap: 14px;
  min-height: 54px;
  padding: 0 18px;
  background: var(--chrome);
  border-bottom: 1px solid var(--chrome-border);
  box-shadow: 0 1px 0 rgb(15 23 42 / 4%);
}

.brand-button {
  padding: 0;
  border: 0;
  background: transparent;
  color: #1e3f78;
  font-size: 15px;
  font-weight: 700;
  text-align: left;
}

.button {
  min-height: 34px;
  padding: 0 12px;
  border: 1px solid #c4cad5;
  border-radius: 6px;
  background: #ffffff;
  color: #202938;
  white-space: nowrap;
}

.button:disabled {
  cursor: not-allowed;
  opacity: .55;
}

.primary-button {
  border-color: #2f5fbd;
  background: #2f5fbd;
  color: #ffffff;
}

.project-select {
  width: 100%;
  min-width: 0;
  height: 34px;
  padding: 0 34px 0 10px;
  border: 1px solid #c4cad5;
  border-radius: 6px;
  background: #ffffff;
  color: #202938;
}

.preview-mode-select {
  width: 100%;
  min-width: 0;
  height: 34px;
  padding: 0 34px 0 10px;
  border: 1px solid #c4cad5;
  border-radius: 6px;
  background: #ffffff;
  color: #202938;
}

.toolbar-status {
  min-width: 0;
  overflow: hidden;
  color: var(--muted);
  font-size: 13px;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.workspace {
  overflow: auto;
  background:
    linear-gradient(90deg, rgb(0 0 0 / 4%) 0 1px, transparent 1px) 0 0 / 24px 24px,
    var(--workspace);
}

.home-section {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 340px;
  gap: 24px;
  max-width: 1180px;
  margin: 0 auto;
  padding: 38px 42px 80px;
}

.home-header {
  grid-column: 1 / -1;
  display: flex;
  align-items: end;
  justify-content: space-between;
  gap: 24px;
}

.home-header h1,
.new-document-form h2,
.document-card h2 {
  margin: 0;
  color: #111827;
}

.home-header h1 {
  font-size: 24px;
  line-height: 1.2;
}

.config-line,
.document-card p,
.form-message {
  margin: 0;
  color: var(--muted);
  font-size: 13px;
}

.document-list {
  display: grid;
  align-content: start;
  gap: 10px;
}

.document-card,
.new-document-form {
  display: grid;
  gap: 12px;
  padding: 16px;
  border: 1px solid #ccd3df;
  border-radius: 8px;
  background: #ffffff;
  box-shadow: 0 8px 24px rgb(15 23 42 / 8%);
}

.document-card {
  grid-template-columns: minmax(0, 1fr) auto;
  align-items: center;
}

.document-card h2,
.document-card p {
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.document-card p {
  grid-column: 1;
}

.document-card .button {
  grid-column: 2;
  grid-row: 1 / span 2;
}

.new-document-form {
  align-content: start;
}

.new-document-form label {
  display: grid;
  gap: 6px;
  color: #384252;
  font-size: 13px;
}

.new-document-form input {
  width: 100%;
  min-width: 0;
  height: 34px;
  padding: 0 10px;
  border: 1px solid #c4cad5;
  border-radius: 6px;
  background: #ffffff;
  color: #202938;
}

.docx-preview-stage {
  min-width: max(100%, 1040px);
  min-height: 100%;
  padding: 0;
}

.docx-preview-frame {
  display: block;
  width: 100%;
  min-width: 840px;
  height: calc(100vh - 54px);
  border: 0;
  background: var(--workspace);
}

.external-preview-message {
  max-width: 640px;
  margin: 40px;
  padding: 18px 20px;
  border: 1px solid #ccd3df;
  border-radius: 8px;
  background: #ffffff;
  box-shadow: 0 8px 24px rgb(15 23 42 / 10%);
}

.external-preview-message h2 {
  margin: 0 0 8px;
  color: #111827;
  font-size: 18px;
  line-height: 1.25;
}

.external-preview-message p {
  margin: 0;
  color: var(--muted);
  font-size: 13px;
  line-height: 1.5;
}

.empty-workspace {
  display: grid;
  place-items: center;
}

.empty-state,
.error-panel {
  max-width: 620px;
  margin: 40px;
  padding: 18px 20px;
  border: 1px solid #ccd3df;
  border-radius: 8px;
  background: #ffffff;
  box-shadow: 0 8px 24px rgb(15 23 42 / 10%);
}

.empty-state {
  color: var(--muted);
}

.error-panel {
  overflow: auto;
  color: #9f1d1d;
  font-family: "Cascadia Code", Consolas, monospace;
  font-size: 12px;
  white-space: pre-wrap;
}

@media (max-width: 980px) {
  .home-section {
    grid-template-columns: 1fr;
  }

  .preview-toolbar {
    grid-template-columns: auto auto minmax(200px, 1fr);
  }
}
`;
}

function greetingNodeTemplate(): string {
  return `/** @jsxImportSource docxcelerate/template */
import { Paragraph, useState } from "docxcelerate/template";
import type { DocumentData } from "../types.ts";

export const Greeting: Paragraph = () => {
  const [state] = useState((data: DocumentData) => ({
    name: data.recipientName,
  }));

  return <Paragraph id="greeting">Hello {state.name},</Paragraph>;
};
`;
}

function introNodeTemplate(): string {
  return `/** @jsxImportSource docxcelerate/template */
import { Paragraph, useState } from "docxcelerate/template";
import type { DocumentData } from "../types.ts";

export const Intro: Paragraph = () => {
  const [state] = useState((data: DocumentData) => ({
    city: data.city,
  }));

  return (
    <Paragraph id="intro">
      We are writing to share an update for {state.city}.
    </Paragraph>
  );
};
`;
}

function sampleBalanceSummaryNodeTemplate(): string {
  return `/** @jsxImportSource docxcelerate/template */
import { dataRef, derive, Paragraph, useState } from "docxcelerate/document";
import type { DocumentData } from "../types.ts";

/**
 * A figure formatted per recipient rather than per build.
 *
 * \`useFormat\` would format it here, which is right whenever the value is known
 * now. This one is not: a published document is written for people whose
 * balances nobody has looked up yet. So the formatting is a deriver the engine
 * runs per document, and the text refers to what it produced.
 */
export const BalanceSummary: Paragraph = () => {
  const [state] = useState((data: DocumentData) => ({
    city: data.city,
  }));

  return (
    <Paragraph
      id="balance-summary"
      derivers={[
        derive("currencyLabel", {
          output: "balanceDueLabel",
          inputs: [dataRef("balanceDue")],
        }),
      ]}
    >
      Your current balance for {state.city} is {"{{derived.balanceDueLabel}}"}.
    </Paragraph>
  );
};
`;
}

function nodesIndexTemplate(): string {
  return `export { Greeting } from "./greeting.node.tsx";
export { Intro } from "./intro.node.tsx";
`;
}

function deriversIndexTemplate(): string {
  return `import type { DeriverDefinitions } from "docxcelerate/document";

export const derivers = {
} satisfies DeriverDefinitions;

export default derivers;
`;
}

function sampleDeriversIndexTemplate(): string {
  return `import type { DeriverDefinitions } from "docxcelerate/document";

export const derivers = {
  currencyLabel: ([amount]) =>
    new Intl.NumberFormat("en-US", {
      style: "currency",
      currency: "USD",
    }).format(Number(amount ?? 0)),
} satisfies DeriverDefinitions;

export default derivers;
`;
}

function documentTemplate(options: { id: string; title: string }): string {
  return `/** @jsxImportSource docxcelerate/template */
import { Document, Section, template } from "docxcelerate/template";
import * as Nodes from "./nodes/index.ts";
import type { DocumentData } from "./types.ts";

export const documentTemplate = template<DocumentData>(
  <Document id="${options.id}" title="${options.title}">
    <Section id="opening" title="Opening">
      <Nodes.Greeting />
      <Nodes.Intro />
    </Section>
  </Document>,
);
`;
}

function sampleDocumentTemplate(): string {
  return `/** @jsxImportSource docxcelerate/template */
import { Document, Section, template } from "docxcelerate/template";
import * as Nodes from "./nodes/index.ts";
import type { DocumentData } from "./types.ts";

export const documentTemplate = template<DocumentData>(
  <Document id="welcome" title="Welcome">
    <Section id="opening" title="Opening">
      <Nodes.Greeting />
      <Nodes.BalanceSummary />
    </Section>
  </Document>,
);
`;
}

function projectTemplate(options: { id: string; title: string }): string {
  return `import { defineDocumentProject } from "docxcelerate/document";
import { derivers } from "./derivers/index.ts";
import { documentTemplate } from "./document.tsx";
import { documentStyle } from "./document-style.ts";
import { previewData } from "./preview-data.ts";
import type { DocumentData } from "./types.ts";

export default defineDocumentProject<DocumentData>({
  id: "${options.id}",
  name: "${options.title}",
  version: "0.1.0",
  template: documentTemplate,
  previewData,
  derivers,
  style: documentStyle,
  previewOptions: {
    availableTokens: 800,
  },
});
`;
}

function sampleNodesIndexTemplate(): string {
  return `export { BalanceSummary } from "./balance-summary.node.tsx";
export { Greeting } from "./greeting.node.tsx";
`;
}

function staticParagraphNodeTemplate(options: { componentName: string; nodeId: string }): string {
  return `/** @jsxImportSource docxcelerate/template */
import { Paragraph, useState } from "docxcelerate/template";
import type { DocumentData } from "../types.ts";

export const ${options.componentName}: Paragraph = () => {
  const [state] = useState((data: DocumentData) => ({
    name: data.recipientName,
  }));

  return (
    <Paragraph id="${options.nodeId}">
      Add ${titleFromSlug(options.nodeId).toLowerCase()} content for {state.name}.
    </Paragraph>
  );
};
`;
}

function dynamicParagraphNodeTemplate(options: { componentName: string; nodeId: string }): string {
  return `/** @jsxImportSource docxcelerate/template */
import {
  Paragraph,
  useAvailableTokens,
  useSetPlaceholders,
  useSetPrompts,
  useState,
} from "docxcelerate/template";
import type { DocumentData } from "../types.ts";

export const ${options.componentName}: Paragraph = () => {
  const availableTokens = useAvailableTokens();
  const [state] = useState((data: DocumentData) => ({
    name: data.recipientName,
  }));

  useSetPrompts({
    generalPrompt:
      \`Write a concise paragraph for \${state.name}. Stay within \${availableTokens} tokens.\`,
    negativePrompt: "Do not invent facts or mention internal implementation details.",
  });

  useSetPlaceholders(
    \`Placeholder ${titleFromSlug(options.nodeId).toLowerCase()} content for \${state.name}.\`,
  );

  return <Paragraph id="${options.nodeId}" />;
};
`;
}

function staticImageNodeTemplate(options: { componentName: string; nodeId: string }): string {
  return `/** @jsxImportSource docxcelerate/template */
import { Image, useState } from "docxcelerate/template";
import type { DocumentData } from "../types.ts";

export const ${options.componentName}: Image = () => {
  const [state] = useState((data: DocumentData) => ({
    name: data.recipientName,
  }));

  return (
    <Image
      id="${options.nodeId}"
      src="assets/${options.nodeId}.png"
      alt={\`${titleFromSlug(options.nodeId)} image for \${state.name}.\`}
    />
  );
};
`;
}

function dynamicImageNodeTemplate(options: { componentName: string; nodeId: string }): string {
  return `/** @jsxImportSource docxcelerate/template */
import {
  Image,
  useAvailableTokens,
  useSetPlaceholders,
  useSetPrompts,
  useState,
} from "docxcelerate/template";
import type { DocumentData } from "../types.ts";

export const ${options.componentName}: Image = () => {
  const availableTokens = useAvailableTokens();
  const [state] = useState((data: DocumentData) => ({
    name: data.recipientName,
  }));

  useSetPrompts({
    generalPrompt:
      \`Describe the image needed for \${state.name}. Stay within \${availableTokens} tokens.\`,
    negativePrompt: "Do not invent facts or include private implementation details.",
  });

  useSetPlaceholders(
    \`Placeholder ${titleFromSlug(options.nodeId).toLowerCase()} image for \${state.name}.\`,
  );

  return <Image id="${options.nodeId}" />;
};
`;
}

function staticGraphNodeTemplate(options: { componentName: string; nodeId: string }): string {
  return `/** @jsxImportSource docxcelerate/template */
import { Graph, useState } from "docxcelerate/template";
import type { DocumentData } from "../types.ts";

export const ${options.componentName}: Graph = () => {
  const [state] = useState((data: DocumentData) => ({
    name: data.recipientName,
  }));

  return (
    <Graph
      id="${options.nodeId}"
      graphType="bar"
      data={{
        labels: ["Current"],
        series: [{ label: state.name, values: [1] }],
      }}
      caption={\`${titleFromSlug(options.nodeId)} for \${state.name}.\`}
    />
  );
};
`;
}

function dynamicGraphNodeTemplate(options: { componentName: string; nodeId: string }): string {
  return `/** @jsxImportSource docxcelerate/template */
import {
  Graph,
  useAvailableTokens,
  useSetPlaceholders,
  useSetPrompts,
  useState,
} from "docxcelerate/template";
import type { DocumentData } from "../types.ts";

export const ${options.componentName}: Graph = () => {
  const availableTokens = useAvailableTokens();
  const [state] = useState((data: DocumentData) => ({
    name: data.recipientName,
  }));

  useSetPrompts({
    generalPrompt:
      \`Prepare graph data for \${state.name}. Stay within \${availableTokens} tokens.\`,
    negativePrompt: "Do not invent facts or include unsupported data points.",
  });

  useSetPlaceholders(
    \`Placeholder ${titleFromSlug(options.nodeId).toLowerCase()} graph for \${state.name}.\`,
  );

  return <Graph id="${options.nodeId}" graphType="bar" />;
};
`;
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

function parentPath(path: string): string {
  const normalized = path.replaceAll("\\", "/");
  const index = normalized.lastIndexOf("/");

  return index === -1 ? "." : normalized.slice(0, index);
}
