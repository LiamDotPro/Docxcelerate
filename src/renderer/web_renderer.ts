import type { DocumentModel, DocumentNode } from "../domain/types.ts";

export interface WebRenderOptions {
  liveReload?: boolean;
  title?: string;
}

export function renderDocumentWebsite(
  letter: DocumentModel,
  options: WebRenderOptions = {},
): string {
  const title = options.title ?? letter.title;

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>${escapeHtml(title)}</title>
    <style>
      :root {
        color-scheme: light;
        font-family: "Aptos", "Segoe UI", Arial, sans-serif;
        background: #f3f4f7;
        color: #1f2933;
        --chrome: #fbfbfd;
        --chrome-border: #d8dce3;
        --workspace: #e6e8ed;
        --page: #ffffff;
        --page-border: #c7ccd6;
        --text-muted: #657287;
        --accent: #2f5fbd;
      }

      * {
        box-sizing: border-box;
      }

      html,
      body {
        margin: 0;
        min-height: 100vh;
      }

      body {
        overflow: hidden;
        background: var(--workspace);
      }

      .word-shell {
        display: grid;
        grid-template-rows: auto auto 1fr;
        min-height: 100vh;
      }

      .titlebar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        min-height: 44px;
        padding: 0 18px;
        background: #204c91;
        color: #ffffff;
      }

      .titlebar strong {
        font-size: 14px;
        font-weight: 600;
      }

      .titlebar-meta {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        font-size: 12px;
        opacity: .94;
      }

      .titlebar-meta::before {
        content: "";
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background: ${options.liveReload ? "#44c767" : "#a9bad8"};
      }

      .ribbon {
        background: var(--chrome);
        border-bottom: 1px solid var(--chrome-border);
        box-shadow: 0 1px 0 rgb(15 23 42 / 4%);
      }

      .tabs {
        display: flex;
        gap: 2px;
        min-height: 34px;
        padding: 0 16px;
        border-bottom: 1px solid var(--chrome-border);
      }

      .tab {
        display: inline-flex;
        align-items: center;
        padding: 0 14px;
        color: #384252;
        font-size: 13px;
      }

      .tab-active {
        border-bottom: 2px solid var(--accent);
        color: var(--accent);
        font-weight: 600;
      }

      .tools {
        display: flex;
        align-items: center;
        gap: 22px;
        min-height: 74px;
        padding: 10px 18px 12px;
      }

      .tool-group {
        display: grid;
        grid-template-columns: repeat(3, 30px);
        gap: 6px;
        padding-right: 20px;
        border-right: 1px solid var(--chrome-border);
      }

      .tool {
        height: 26px;
        border: 1px solid #d4d9e2;
        border-radius: 4px;
        background: linear-gradient(#ffffff, #f0f2f6);
      }

      .tool-wide {
        width: 116px;
        height: 26px;
        border: 1px solid #d4d9e2;
        border-radius: 4px;
        background: #ffffff;
      }

      .workspace {
        overflow: auto;
        background:
          linear-gradient(90deg, rgb(0 0 0 / 4%) 0 1px, transparent 1px) 0 0 / 24px 24px,
          var(--workspace);
      }

      .workspace-inner {
        width: max-content;
        min-width: 1180px;
        min-height: calc(100vh - 153px);
        padding: 0 96px 80px;
      }

      .horizontal-ruler {
        position: sticky;
        top: 0;
        z-index: 1;
        width: 210mm;
        height: 28px;
        margin: 0 auto 18px;
        background:
          repeating-linear-gradient(
            90deg,
            #f8f9fb 0,
            #f8f9fb 9.5mm,
            #c7ccd6 9.5mm,
            #c7ccd6 10mm
          );
        border: 1px solid #cfd4dd;
        border-top: 0;
        box-shadow: 0 1px 2px rgb(15 23 42 / 10%);
      }

      .page-stage {
        display: grid;
        place-items: start center;
        min-width: 100%;
      }

      .a4-page {
        width: 210mm;
        min-height: 297mm;
        padding: 25.4mm;
        background: var(--page);
        border: 1px solid var(--page-border);
        box-shadow:
          0 1px 2px rgb(15 23 42 / 12%),
          0 16px 44px rgb(15 23 42 / 24%);
      }

      .document-title {
        margin: 0 0 18pt;
        font-family: Cambria, Georgia, serif;
        font-size: 20pt;
        line-height: 1.2;
        font-weight: 700;
        color: #111827;
      }

      .letter-section {
        margin: 0 0 12pt;
      }

      .letter-section h2 {
        margin: 16pt 0 7pt;
        font-size: 12pt;
        line-height: 1.3;
        font-weight: 700;
        color: #111827;
      }

      .letter-paragraph {
        margin: 0 0 10pt;
        font-family: Calibri, "Aptos", "Segoe UI", Arial, sans-serif;
        font-size: 11pt;
        line-height: 1.42;
        color: #111827;
      }

      .node-id {
        display: block;
        margin-bottom: 3pt;
        color: var(--text-muted);
        font-size: 8pt;
        line-height: 1.2;
        text-transform: uppercase;
      }

      .image-placeholder,
      .graph-placeholder,
      .toc-placeholder {
        margin: 10pt 0;
        padding: 9pt;
        border: 1px dashed #9aa6b8;
        color: #59677a;
        font-size: 9pt;
      }

      @media (max-width: 900px) {
        body {
          overflow: auto;
        }

        .word-shell {
          min-width: 900px;
        }
      }
    </style>
  </head>
  <body>
    <main class="word-shell">
      <header class="titlebar">
        <strong>${escapeHtml(letter.title)}</strong>
        <span class="titlebar-meta">${
    options.liveReload ? "Live preview connected" : "Preview"
  }</span>
      </header>

      <section class="ribbon" aria-label="Document toolbar">
        <nav class="tabs" aria-label="Ribbon tabs">
          <span class="tab tab-active">Home</span>
          <span class="tab">Insert</span>
          <span class="tab">Layout</span>
          <span class="tab">Review</span>
          <span class="tab">View</span>
        </nav>
        <div class="tools" aria-hidden="true">
          <div class="tool-wide"></div>
          <div class="tool-group">
            <span class="tool"></span>
            <span class="tool"></span>
            <span class="tool"></span>
            <span class="tool"></span>
            <span class="tool"></span>
            <span class="tool"></span>
          </div>
          <div class="tool-wide"></div>
          <div class="tool-wide"></div>
        </div>
      </section>

      <section class="workspace" aria-label="Document workspace">
        <div class="workspace-inner">
          <div class="horizontal-ruler" aria-hidden="true"></div>
          <div class="page-stage">
            <article class="a4-page" aria-label="${escapeHtml(letter.title)}">
              <h1 class="document-title">${escapeHtml(letter.title)}</h1>
              ${letter.nodes.map(renderNode).join("\n")}
            </article>
          </div>
        </div>
      </section>
    </main>
    ${options.liveReload ? liveReloadScript() : ""}
  </body>
</html>`;
}

function renderNode(node: DocumentNode): string {
  if (node.kind === "section") {
    return `<section class="letter-section" data-node-id="${escapeHtml(node.id)}">
      <h2>${escapeHtml(node.title ?? node.id)}</h2>
      ${node.children.map(renderNode).join("\n")}
    </section>`;
  }

  if (node.kind === "paragraph") {
    const className = node.mode === "dynamic"
      ? "letter-paragraph paragraph-dynamic"
      : "letter-paragraph paragraph-static";

    return `<p class="${className}" data-node-id="${escapeHtml(node.id)}">
      ${
      node.mode === "dynamic" ? `<span class="node-id">Dynamic: ${escapeHtml(node.id)}</span>` : ""
    }
      ${escapeHtml(node.text ?? "")}
    </p>`;
  }

  if (node.kind === "image") {
    return `<div class="image-placeholder" data-node-id="${escapeHtml(node.id)}">
      Image: ${escapeHtml(node.alt ?? node.path ?? node.placeholder ?? node.id)}
    </div>`;
  }

  if (node.kind === "graph") {
    return `<div class="graph-placeholder" data-node-id="${escapeHtml(node.id)}">
      ${escapeHtml(node.graphType)} graph: ${
      escapeHtml(node.caption ?? node.placeholder ?? node.id)
    }
    </div>`;
  }

  return `<div class="toc-placeholder" data-node-id="${escapeHtml(node.id)}">
    ${escapeHtml(node.title ?? "Table of contents")}
  </div>`;
}

function liveReloadScript(): string {
  return `<script>
    const events = new EventSource("/events");
    events.addEventListener("reload", () => window.location.reload());
  </script>`;
}

function escapeHtml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
