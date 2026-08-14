import { paragraph } from "docxcelerate";
import type { RepairsData } from "../types.ts";

export const Appointment = paragraph<RepairsData>({
  id: "appointment",
  render: (data) =>
    `We have booked a ${data.trade} to visit ${data.address} on ` +
    `${data.visitDate}, ${data.visitWindow}. Your job reference is ` +
    `${data.jobRef}.`,
});
