import { Paragraph, useState } from "docxcelerate/template";
import type { OfferData } from "../types.ts";

export const Offer: Paragraph = () => {
  const [state] = useState((data: OfferData) => ({
    programme: data.programme,
    college: data.college,
    startDate: data.startDate,
    offerRef: data.offerRef,
  }));

  return (
    <Paragraph id="offer">
      I am delighted to offer you a place on the {state.programme} at {state.college},
      beginning {state.startDate}. Your application reference is {state.offerRef}; please
      quote it in any correspondence with us.
    </Paragraph>
  );
};
