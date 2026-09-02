import { buildDocument, createDocumentProjectArtifact } from "docxcelerate";
import type { DocumentModel, DocumentProject } from "docxcelerate/document";
import { settleDocxPreview } from "docxcelerate/preview";
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

type UploadBodyKind =
  | "artifact"
  | "document"
  | "letter"
  | "stored-document"
  | "stored-letter";

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
  model: DocumentModel,
  renderer: PreviewRenderer,
): Promise<HTMLElement> {
  const { createDocxBlob } = await import("docxcelerate/docx");
  const documentBlob = await createDocxBlob(model);

  if (renderer === "docx-preview") {
    return await renderClientDocxPreview(model, documentBlob);
  }

  return await renderHostedDocxPreview(model, documentBlob, renderer);
}

async function renderClientDocxPreview(model: DocumentModel, documentBlob: Blob): Promise<HTMLElement> {
  const docxPreview = await import("docx-preview");
  const stage = document.createElement("div");
  stage.className = "docx-preview-stage";

  const frame = document.createElement("iframe");
  frame.className = "docx-preview-frame";
  frame.title = model.title + " DOCX preview";

  const body = document.createElement("body");
  const head = document.createElement("head");

  await docxPreview.renderAsync(documentBlob, body, head, {
    inWrapper: true,
    breakPages: true,
    ignoreLastRenderedPageBreak: false,
  });

  // docx-preview does not read everything the file says: it drops a field run,
  // looks for a table's indent under an attribute that never carries it, and
  // wraps a picture in an element a paragraph may not hold. This finishes the
  // reading, so what is shown is what Word will show.
  settleDocxPreview(body, model);

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
  model: DocumentModel,
  documentBlob: Blob,
  renderer: Exclude<PreviewRenderer, "docx-preview">,
): Promise<HTMLElement> {
  const stage = document.createElement("div");
  stage.className = "docx-preview-stage";

  if (isPrivatePreviewHost()) {
    stage.append(renderHostedPreviewNotice(renderer));
    return stage;
  }

  const documentUrl = await uploadPreviewDocx(model, documentBlob);
  const frame = document.createElement("iframe");
  frame.className = "docx-preview-frame";
  frame.title = model.title + " " + previewRendererLabel(renderer);
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

async function uploadPreviewDocx(model: DocumentModel, documentBlob: Blob): Promise<string> {
  const response = await fetch(
    "/api/docxcelerate/preview-docx?name=" + encodeURIComponent(model.id + ".docx"),
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
  ].join("\n");
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
  return path.replace(/^\.\.\//, "");
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
