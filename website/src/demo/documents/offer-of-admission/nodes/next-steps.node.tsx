import { Paragraph, useState } from "docxcelerate/template";
import type { OfferData } from "../types.ts";

export const NextSteps: Paragraph = () => {
  const [state] = useState((data: OfferData) => ({
    replyBy: data.replyBy,
    college: data.college,
  }));

  return (
    <Paragraph id="how-to-accept">
      To accept, reply through the applicant portal by {state.replyBy}. If you would like
      to visit {state.college} before deciding, our open afternoons run every Thursday
      through April and you are very welcome.
    </Paragraph>
  );
};
