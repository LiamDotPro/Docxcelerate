import { image } from "docxcelerate";
import type { SampleData } from "../sample-data.ts";

/**
 * Every field takes a plain value or a function of your data, so a signature
 * that varies by manager needs no branching in the template.
 */
export const Signature = image<SampleData>({
  id: "signature",
  src: (data) => data.signatureUrl,
  alt: (data) => `Signed by ${data.managerName}`,
  width: 180,
  height: 60,
});
