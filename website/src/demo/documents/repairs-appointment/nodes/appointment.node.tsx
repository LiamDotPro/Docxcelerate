/** @jsxImportSource docxcelerate/template */
import { Paragraph, useState } from "docxcelerate/template";
import type { RepairsData } from "../types.ts";

export const Appointment: Paragraph = () => {
  const [state] = useState((data: RepairsData) => ({
    trade: data.trade,
    address: data.address,
    visitDate: data.visitDate,
    visitWindow: data.visitWindow,
    jobRef: data.jobRef,
  }));

  return (
    <Paragraph id="appointment-details">
      We have booked a {state.trade} to visit {state.address} on {state.visitDate},{" "}
      {state.visitWindow}. Your job reference is {state.jobRef}.
    </Paragraph>
  );
};
