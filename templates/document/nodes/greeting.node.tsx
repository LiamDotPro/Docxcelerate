import { Paragraph, useState } from "docxcelerate/template";
import type { DocumentData } from "../types.ts";

export const Greeting: Paragraph = () => {
  const [state] = useState((data: DocumentData) => ({
    name: data.recipientName,
  }));

  return <Paragraph id="greeting">Hello {state.name},</Paragraph>;
};
