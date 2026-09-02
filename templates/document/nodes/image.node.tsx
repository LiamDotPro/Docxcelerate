import { Image, useState } from "docxcelerate/template";
import type { DocumentData } from "../types.ts";

export const __NODE_COMPONENT__: Image = () => {
  const [state] = useState((data: DocumentData) => ({
    name: data.recipientName,
  }));

  return (
    <Image
      id="__NODE_ID__"
      src="assets/__NODE_ID__.png"
      alt={`__NODE_TITLE__ image for ${state.name}.`}
    />
  );
};
