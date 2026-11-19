import { defineDocumentProject } from "docxcelerate/document";
import { documentTemplate } from "./document.tsx";
import { previewData } from "./preview-data.ts";
import { reviewStyle } from "./review-style.ts";
import type { PortfolioReviewData } from "./types.ts";

export default defineDocumentProject<PortfolioReviewData>({
  id: "portfolio-review",
  name: "Quarterly portfolio review",
  version: "1.0.0",
  template: documentTemplate,
  style: reviewStyle,
  previewData,
});
