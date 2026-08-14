import { paragraph } from "docxcelerate";
import type { RepairsData } from "../types.ts";

/**
 * An absent optional field should change the sentence, not print "undefined"
 * or leave a blank line.
 */
export const Access = paragraph<RepairsData>({
  id: "access",
  render: (data) =>
    data.accessNotes
      ? `We hold the following access note for your home: ${data.accessNotes} ` +
        `Please let us know if this is out of date.`
      : `Someone aged 18 or over needs to be home for the visit. If that is ` +
        `not possible, call us and we will arrange access another way.`,
});
