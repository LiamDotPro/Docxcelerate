import { cleanMinimalDocumentStyle } from "docxcelerate";
import { defineDocumentProject } from "docxcelerate/document";
import { letterTemplate } from "./letter.tsx";
import { previewData } from "./preview-data.ts";
import type { OfferData } from "./types.ts";

export default defineDocumentProject<OfferData>({
  id: "offer-of-admission",
  name: "Offer of Admission",
  version: "1.0.0",
  template: letterTemplate,
  style: cleanMinimalDocumentStyle,
  previewData,
});
