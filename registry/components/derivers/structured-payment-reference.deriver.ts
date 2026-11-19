import { deriver } from "docxcelerate/template";

/**
 * Works out, validates and groups a structured payment reference, or decides
 * that the invoice number has to be quoted instead. Installed by
 * `dxcl add structured-payment-reference`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level helper or constant would be a ReferenceError there. Every
 * check is a loop over digit characters; no Number or BigInt is made from a
 * base, because that would drop leading zeros and lose precision past 2^53.
 */
export default deriver({
  name: "structuredPaymentReference",
  run: (payment: unknown, invoice: unknown) => {
    // deno-lint-ignore no-explicit-any
    const pay = payment as any;
    // deno-lint-ignore no-explicit-any
    const inv = invoice as any;

    if (typeof inv?.number !== "string" || inv.number.trim() === "") {
      throw new Error("invoice.number is required");
    }
    const fallback = "Please quote invoice number " + inv.number.trim() + " with your payment.";
    const invalid = { issued: null, invalid: fallback, unissuable: null };
    const unissuable = { issued: null, invalid: null, unissuable: fallback };
    if (pay === undefined || pay === null) return invalid;

    const schemes = ["rf", "no-kid-mod10", "no-kid-mod11", "fi", "fi-rf", "be-ogm", "ch-qrr", "se-ocr", "se-ocr-length"];
    if (!schemes.includes(pay.scheme)) {
      throw new Error("payment.scheme must be one of " + schemes.join(", "));
    }
    const present = (v: unknown) => v !== undefined && v !== null;
    if (present(pay.reference) && typeof pay.reference !== "string") {
      throw new Error("payment.reference must be a string");
    }
    if (present(pay.base) && typeof pay.base !== "string") {
      throw new Error("payment.base must be a string (a number loses leading zeros)");
    }
    const scheme: string = pay.scheme;

    const strip = (s: string) => s.replace(/\s/g, "");
    const reference = present(pay.reference) ? strip(pay.reference) : "";
    const base = present(pay.base) ? strip(pay.base) : "";
    const validating = reference !== "";
    if (!validating && base === "") return invalid;
    let input = validating ? reference : base;
    if (scheme === "be-ogm") input = input.replace(/[+*/]/g, "");

    const digitsOnly = (s: string) => /^\d*$/.test(s);
    const between = (s: string, min: number, max: number) => s.length >= min && s.length <= max;

    const mod97 = (s: string) => {
      let r = 0;
      for (const ch of s) {
        const v = /\d/.test(ch) ? ch : String(ch.charCodeAt(0) - 55);
        for (const d of v) r = (r * 10 + Number(d)) % 97;
      }
      return r;
    };
    const rfCheck = (body: string) => String(98 - mod97(body + "RF00")).padStart(2, "0");
    const rfValid = (ref: string) => /^RF\d{2}[A-Z0-9]{1,21}$/.test(ref) && mod97(ref.slice(4) + ref.slice(0, 4)) === 1;
    const luhn = (p: string) => {
      let sum = 0;
      for (let i = 0; i < p.length; i++) {
        let d = Number(p[p.length - 1 - i]);
        if (i % 2 === 0) {
          d *= 2;
          if (d > 9) d -= 9;
        }
        sum += d;
      }
      return String((10 - sum % 10) % 10);
    };
    const mod11 = (p: string) => {
      let sum = 0;
      for (let i = 0; i < p.length; i++) sum += Number(p[p.length - 1 - i]) * (2 + i % 6);
      const c = 11 - sum % 11;
      return c === 11 ? "0" : c === 10 ? "-" : String(c);
    };
    const fi731 = (p: string) => {
      const w = [7, 3, 1];
      let sum = 0;
      for (let i = 0; i < p.length; i++) sum += Number(p[p.length - 1 - i]) * w[i % 3];
      return String((10 - sum % 10) % 10);
    };
    const qrr = (p: string) => {
      const table = [0, 9, 4, 6, 8, 2, 7, 1, 3, 5];
      let carry = 0;
      for (const ch of p) carry = table[(carry + Number(ch)) % 10];
      return String((10 - carry) % 10);
    };
    const ogm = (b10: string) => String(mod97(b10) || 97).padStart(2, "0");
    const chunk = (s: string, n: number) => s.match(new RegExp(".{1," + n + "}", "g"))!.join(" ");
    const chunkRight = (s: string, n: number) => {
      const parts: string[] = [];
      for (let end = s.length; end > 0; end -= n) parts.unshift(s.slice(Math.max(0, end - n), end));
      return parts.join(" ");
    };
    const stripZeros = (s: string) => s.replace(/^0+/, "");

    // Each scheme gives the finished reference (grouped), or null for invalid.
    // A fi reference without its grouping is also what fi-rf builds on.
    const fiNational = (s: string): string | null => {
      if (!digitsOnly(s)) return null;
      const n = stripZeros(s);
      if (validating) return between(n, 4, 20) && n[n.length - 1] === fi731(n.slice(0, -1)) ? n : null;
      return between(n, 3, 19) ? n + fi731(n) : null;
    };
    let label = "";
    let grouped: string | null = null;
    let notIssuable = false;

    if (scheme === "rf" || scheme === "fi-rf") {
      label = "Creditor reference (RF)";
      const up = input.toUpperCase();
      let ref: string | null = null;
      if (scheme === "rf") {
        if (validating) ref = rfValid(up) ? up : null;
        else if (/^[A-Z0-9]{1,21}$/.test(up)) ref = "RF" + rfCheck(up) + up;
      } else if (validating) {
        const body = up.slice(4);
        ref = rfValid(up) && digitsOnly(body) && body[0] !== "0" && fiNational(body) !== null ? up : null;
      } else {
        const national = fiNational(up);
        if (national !== null) ref = "RF" + rfCheck(national) + national;
      }
      if (ref !== null) grouped = chunk(ref, 4);
    } else if (scheme === "no-kid-mod10" || scheme === "no-kid-mod11") {
      label = "KID";
      const is11 = scheme === "no-kid-mod11";
      const check = is11 ? mod11 : luhn;
      if (validating) {
        const body = input.slice(0, -1);
        const last = input[input.length - 1];
        if (between(input, 2, 25) && digitsOnly(body) && (digitsOnly(last) || (is11 && last === "-"))) {
          const c = check(body);
          if (last === c) {
            if (c === "-") notIssuable = true;
            else grouped = input;
          }
        }
      } else if (digitsOnly(input) && between(input, 1, 24)) {
        const c = check(input);
        if (c === "-") notIssuable = true;
        else grouped = input + c;
      }
    } else if (scheme === "fi") {
      label = "Reference number (viitenumero)";
      const n = fiNational(input);
      if (n !== null) grouped = chunkRight(n, 5);
    } else if (scheme === "be-ogm") {
      label = "Structured communication (OGM/VCS)";
      if (digitsOnly(input)) {
        let d: string | null = null;
        if (validating) {
          if (input.length === 12 && input.slice(10) === ogm(input.slice(0, 10))) d = input;
        } else if (between(input, 1, 10)) {
          const b = input.padStart(10, "0");
          d = b + ogm(b);
        }
        if (d !== null) grouped = "+++" + d.slice(0, 3) + "/" + d.slice(3, 7) + "/" + d.slice(7) + "+++";
      }
    } else if (scheme === "ch-qrr") {
      label = "QR reference";
      if (digitsOnly(input)) {
        let d: string | null = null;
        if (validating) {
          if (input.length === 27 && input[26] === qrr(input.slice(0, 26))) d = input;
        } else if (between(input, 1, 26)) {
          const b = input.padStart(26, "0");
          d = b + qrr(b);
        }
        if (d !== null) grouped = d.slice(0, 2) + " " + chunk(d.slice(2), 5);
      }
    } else {
      label = "OCR number";
      if (digitsOnly(input)) {
        let d: string | null = null;
        if (scheme === "se-ocr") {
          if (validating) {
            if (between(input, 2, 25) && input[input.length - 1] === luhn(input.slice(0, -1))) d = input;
          } else if (between(input, 1, 24)) d = input + luhn(input);
        } else if (validating) {
          if (
            between(input, 3, 25) && input[input.length - 1] === luhn(input.slice(0, -1)) &&
            input[input.length - 2] === String(input.length % 10)
          ) d = input;
        } else if (between(input, 1, 23)) {
          const withLength = input + String((input.length + 2) % 10);
          d = withLength + luhn(withLength);
        }
        if (d !== null) grouped = d;
      }
    }

    if (notIssuable) return unissuable;
    if (grouped === null) return invalid;
    return { issued: label + ": " + grouped, invalid: null, unissuable: null };
  },
});
