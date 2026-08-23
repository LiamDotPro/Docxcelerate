import { defineDocumentProject } from "docxcelerate/document";
import { documentTemplate } from "./document.tsx";
import { invoiceStyle } from "./invoice-style.ts";
import { previewData } from "./preview-data.ts";
import type { InvoiceData } from "./types.ts";

export default defineDocumentProject<InvoiceData>({
  id: "invoice",
  name: "Invoice",
  version: "1.0.0",
  template: documentTemplate,
  style: invoiceStyle,
  previewData,
});
