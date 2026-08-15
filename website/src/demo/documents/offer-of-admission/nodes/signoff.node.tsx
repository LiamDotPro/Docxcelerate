import { Paragraph, useState } from "docxcelerate/template";
import type { OfferData } from "../types.ts";

export const SignOff: Paragraph = () => {
  const [state] = useState((data: OfferData) => ({
    name: data.signatory.name,
    title: data.signatory.title,
    college: data.college,
  }));

  return (
    <Paragraph id="sign-off">
      Yours sincerely, {state.name}, {state.title}, {state.college}.
    </Paragraph>
  );
};
