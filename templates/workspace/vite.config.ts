import { randomUUID } from "node:crypto";
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
            /^\/api\/docxcelerate\/preview-docx\/([^/]+)\/[^/]+[.]docx$/,
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
    return dirname(entrypoint.replace(/^[.][\\/]+/, ""));
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
  await writeFile(path, JSON.stringify(value, null, 2) + "\n", "utf8");
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
  response.end(JSON.stringify(value) + "\n");
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
