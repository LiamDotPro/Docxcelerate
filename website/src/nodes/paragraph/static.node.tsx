import { Paragraph, useState } from "docxcelerate/template";
import type { SampleData } from "../sample-data.ts";

/**
 * The smallest useful node: data taken into state, and a line built from it.
 * `useState` is where data enters a component, and the only place it does.
 */
export const Greeting: Paragraph = () => {
  const [state] = useState((data: SampleData) => ({ name: data.memberName }));

  return <Paragraph id="greeting">Dear {state.name},</Paragraph>;
};
