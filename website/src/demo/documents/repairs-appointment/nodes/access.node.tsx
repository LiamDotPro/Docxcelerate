/** @jsxImportSource docxcelerate/template */
import { Paragraph, useState } from "docxcelerate/template";
import type { RepairsData } from "../types.ts";

/**
 * An absent optional field should change the sentence, not print "undefined"
 * or leave a blank line. Two arms, two ids — so the resolved document records
 * which of the two this resident was sent.
 */
export const Access: Paragraph = () => {
  const [state] = useState((data: RepairsData) => ({ notes: data.accessNotes }));

  if (state.notes) {
    return (
      <Paragraph id="access-noted">
        We hold the following access note for your home: {state.notes} Please let us know
        if this is out of date.
      </Paragraph>
    );
  }

  return (
    <Paragraph id="access-default">
      Someone aged 18 or over needs to be home for the visit. If that is not possible,
      call us and we will arrange access another way.
    </Paragraph>
  );
};
