import { Paragraph, useState } from "docxcelerate/template";
import type { DocumentData } from "../types.ts";

export const Intro: Paragraph = () => {
  const [state] = useState((data: DocumentData) => ({
    city: data.city,
  }));

  return (
    <Paragraph id="intro">
      We are writing to share an update for {state.city}.
    </Paragraph>
  );
};
