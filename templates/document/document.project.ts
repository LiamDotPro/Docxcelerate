import { defineDocumentProject } from "docxcelerate/document";
import { derivers } from "./derivers/index.ts";
import { documentTemplate } from "./document.tsx";
import { documentStyle } from "./document-style.ts";
import { previewData } from "./preview-data.ts";
import type { DocumentData } from "./types.ts";

export default defineDocumentProject<DocumentData>({
  id: "__DOCUMENT_ID__",
  name: "__DOCUMENT_TITLE__",
  version: "0.1.0",
  template: documentTemplate,
  previewData,
  derivers,
  style: documentStyle,
  previewOptions: {
    availableTokens: 800,
  },
});
