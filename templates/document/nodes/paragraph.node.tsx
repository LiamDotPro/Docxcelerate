import { Paragraph, useState } from "docxcelerate/template";
import type { DocumentData } from "../types.ts";

export const __NODE_COMPONENT__: Paragraph = () => {
  const [state] = useState((data: DocumentData) => ({
    name: data.recipientName,
  }));

  return (
    <Paragraph id="__NODE_ID__">
      Add __NODE_TITLE_LOWER__ content for {state.name}.
    </Paragraph>
  );
};
