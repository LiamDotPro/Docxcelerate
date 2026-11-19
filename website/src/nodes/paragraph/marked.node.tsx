import { Paragraph, Text, useState } from "docxcelerate/template";
import type { SampleData } from "../sample-data.ts";

/**
 * Some of the words set apart: the date a member has to act by, and a phrase
 * that leads somewhere. `<Text>` says what the words are and never what they
 * look like, so the link takes the theme's accent and re-themes with the page.
 */
export const RenewalTerms: Paragraph = () => {
  const [state] = useState((data: SampleData) => ({
    plan: data.plan,
    renewsOn: data.renewsOn,
  }));

  return (
    <Paragraph id="renewal-terms">
      Your {state.plan} membership renews on <Text bold>{state.renewsOn}</Text>. To change or
      cancel it before then, <Text href="https://example.com/membership">manage it online</Text>.
    </Paragraph>
  );
};
