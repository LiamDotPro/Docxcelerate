import { Paragraph, useState } from "docxcelerate/template";
import type { PolicyData } from "../types.ts";

export const Greeting: Paragraph = () => {
  const [state] = useState((data: PolicyData) => ({ name: data.holderName }));

  return <Paragraph id="greeting">Dear {state.name},</Paragraph>;
};
