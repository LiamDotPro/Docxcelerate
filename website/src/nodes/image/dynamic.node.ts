import { image } from "docxcelerate";
import type { SampleData } from "../sample-data.ts";

/**
 * The same prompt set as a dynamic paragraph, for artwork the endpoint
 * produces rather than something you hold. The placeholder is what previews
 * show, so the layout is settled before any image exists.
 */
export const CentrePhoto = image<SampleData>({
  id: "centre-photo",
  placeholder: (data) => `Photograph of ${data.centreName}`,
  generalPrompt: (data) =>
    `A wide daylight photograph of the entrance to ${data.centreName}, ` +
    `people arriving, no text overlay.`,
  negativePrompt: () => `No logos, no recognisable faces, no stock-photo staging.`,
});
