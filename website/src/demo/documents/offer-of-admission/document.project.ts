import { cleanMinimalDocumentStyle } from "docxcelerate";
import { defineDocumentProject } from "docxcelerate/document";
import { documentTemplate } from "./document.tsx";
import { previewData } from "./preview-data.ts";
import type { OfferData } from "./types.ts";

export default defineDocumentProject<OfferData>({
  id: "offer-of-admission",
  name: "Offer of Admission",
  version: "1.0.0",
  template: documentTemplate,
  style: cleanMinimalDocumentStyle,
  previewData,
});
