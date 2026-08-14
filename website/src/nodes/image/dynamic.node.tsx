/** @jsxImportSource docxcelerate/template */
import { Image, useSetPlaceholders, useSetPrompts, useState } from "docxcelerate/template";
import type { SampleData } from "../sample-data.ts";

/**
 * The same prompt set as a dynamic paragraph, for artwork the engine produces
 * rather than something you hold. The placeholder is what previews show, so
 * the layout is settled before any image exists.
 */
export const CentrePhoto: Image = () => {
  const [state] = useState((data: SampleData) => ({ centreName: data.centreName }));

  useSetPrompts({
    generalPrompt: `A wide daylight photograph of the entrance to ${state.centreName}, ` +
      `people arriving, no text overlay.`,
    negativePrompt: "No logos, no recognisable faces, no stock-photo staging.",
  });

  useSetPlaceholders(`Photograph of ${state.centreName}`);

  return <Image id="centre-photo" />;
};
