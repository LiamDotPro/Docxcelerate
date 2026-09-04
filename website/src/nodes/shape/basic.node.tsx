import { Paragraph, type Shape as ShapeComponent, Shape, useFormat, useState } from "docxcelerate/template";
import type { SampleData } from "../sample-data.ts";

/**
 * A drawn rectangle with the document's own words on it.
 *
 * The height is the point. A banner that grew with its wording would be a
 * different depth on every letter it went out on, and a reader flicking
 * through a stack would read the shorter one as a different kind of notice.
 */
export const RenewalBanner: ShapeComponent = () => {
  const { date } = useFormat("en-GB");
  const [state] = useState((data: SampleData) => ({
    plan: data.plan,
    renewsOn: data.renewsOn,
  }));

  return (
    <Shape id="renewal-banner" variant="banner" height={44}>
      <Paragraph id="renewal-banner-line">
        {state.plan} renews on {date(state.renewsOn)}
      </Paragraph>
    </Shape>
  );
};
