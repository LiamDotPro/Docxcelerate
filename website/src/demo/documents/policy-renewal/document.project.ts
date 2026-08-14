import { cleanMinimalDocumentStyle } from "docxcelerate";
import { defineDocumentProject } from "docxcelerate/document";
import { documentTemplate } from "./document.tsx";
import { previewData } from "./preview-data.ts";
import type { PolicyData } from "./types.ts";

export default defineDocumentProject<PolicyData>({
  id: "policy-renewal",
  name: "Policy Renewal",
  version: "0.4.2",
  template: documentTemplate,
  style: cleanMinimalDocumentStyle,
  previewData,
});
