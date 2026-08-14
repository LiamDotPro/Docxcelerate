import { cleanMinimalDocumentStyle } from "docxcelerate";
import { defineDocumentProject } from "docxcelerate/document";
import { documentTemplate } from "./document.tsx";
import { previewData } from "./preview-data.ts";
import type { RepairsData } from "./types.ts";

export default defineDocumentProject<RepairsData>({
  id: "repairs-appointment",
  name: "Repair Appointment",
  version: "2.1.0",
  template: documentTemplate,
  style: cleanMinimalDocumentStyle,
  previewData,
});
