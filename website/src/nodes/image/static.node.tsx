/** @jsxImportSource docxcelerate/template */
import { Image, useState } from "docxcelerate/template";
import type { SampleData } from "../sample-data.ts";

/**
 * Every field is a plain prop, computed from state before the element is
 * returned — so a signature that varies by manager needs no branching in the
 * template that places it.
 */
export const Signature: Image = () => {
  const [state] = useState((data: SampleData) => ({
    src: data.signatureUrl,
    manager: data.managerName,
  }));

  return (
    <Image
      id="signature"
      src={state.src}
      alt={`Signed by ${state.manager}`}
      width={180}
      height={60}
    />
  );
};
