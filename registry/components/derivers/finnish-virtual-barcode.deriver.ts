// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Builds the 54-digit Finnish virtual bank barcode, or says why it must not be
 * printed. Installed by `dxcl add finnish-virtual-barcode`.
 *
 * It never throws: every bad input is a withheld code, because the fallback
 * (the IBAN and reference on their own) is still a payable invoice.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 */
export default deriver({
  name: "finnishVirtualBarcode",
  run: (payee: unknown, invoice: unknown, payment: unknown, virtualBarcode: unknown) => {
    const pay = payee as any;
    const inv = invoice as any;
    const pmt = payment as any;

    const iban = String(pay?.iban ?? "").replace(/\s+/g, "").toUpperCase();
    const ref = String(pmt?.reference ?? "").replace(/\s+/g, "").toUpperCase();
    const erp = String(virtualBarcode ?? "").replace(/\s+/g, "");
    const currency = String(inv?.currency ?? "").trim().toUpperCase();

    // Folded digit by digit: Number() on the whole string loses precision.
    const mod97 = (s: string): boolean => {
      const moved = s.slice(4) + s.slice(0, 4);
      let r = 0;
      for (const ch of moved) {
        const n = /[A-Z]/.test(ch) ? String(ch.charCodeAt(0) - 55) : ch;
        for (const d of n) r = (r * 10 + Number(d)) % 97;
      }
      return r === 1;
    };
    const groups = (s: string, size: number, fromRight: boolean): string => {
      const out: string[] = [];
      if (fromRight) {
        for (let i = s.length; i > 0; i -= size) out.unshift(s.slice(Math.max(0, i - size), i));
      } else {
        for (let i = 0; i < s.length; i += size) out.push(s.slice(i, i + size));
      }
      return out.join(" ");
    };

    const ibanOk = /^[A-Z]{2}\d{2}[A-Z0-9]{1,30}$/.test(iban) && mod97(iban);

    // The reference is judged on its own, so it is grouped even when the IBAN
    // failed first.
    const isRf = ref.startsWith("RF");
    const v = ref.replace(/^0+/, "");
    let refOk: boolean;
    if (isRf) {
      refOk = /^RF\d{2}[0-9A-Z]{1,21}$/.test(ref) && mod97(ref);
    } else {
      refOk = /^\d{4,20}$/.test(v);
      if (refOk) {
        let sum = 0;
        const body = v.slice(0, -1);
        for (let i = 0; i < body.length; i++) {
          sum += Number(body[body.length - 1 - i]) * [7, 3, 1][i % 3];
        }
        refOk = (10 - sum % 10) % 10 === Number(v[v.length - 1]);
      }
    }
    const printedIban = ibanOk ? groups(iban, 4, false) : String(pay?.iban ?? "").trim();
    const printedRef = !refOk
      ? String(pmt?.reference ?? "").trim()
      : isRf
      ? groups(ref, 4, false)
      : groups(v, 5, true);

    const withheld = (reason: string) => ({
      outcome: "withheld",
      reason,
      iban: printedIban,
      reference: printedRef,
      code: "",
    });

    // 1. Nothing owed: a credit note or a paid invoice prints nothing at all.
    const open = inv?.openAmount === true;
    if (!open && typeof inv?.total === "number" && Number.isFinite(inv.total) && inv.total <= 0) {
      return { outcome: "none", reason: "", iban: "", reference: "", code: "" };
    }
    if (!ibanOk) return withheld("iban");
    if (!/^FI\d{16}$/.test(iban)) return withheld("not-finnish");
    if (currency !== "EUR") return withheld("currency");

    let amount8 = "00000000";
    if (!open) {
      const total = inv?.total;
      if (typeof total !== "number" || !Number.isFinite(total)) return withheld("amount");
      const cents = Math.round(total * 100);
      if (Math.abs(total * 100 - cents) > 1e-6) return withheld("amount");
      // Zero-filled past the cap, never truncated: the payer types it in.
      if (cents <= 99_999_999) amount8 = String(cents).padStart(8, "0");
    }

    if (!refOk) return withheld("reference");
    if (isRf && !/^\d+$/.test(ref.slice(4))) return withheld("rf-alphanumeric");

    let due6 = "000000";
    const dueDate = inv?.dueDate;
    if (inv?.onReceipt !== true && dueDate !== undefined && dueDate !== null && dueDate !== "") {
      const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(dueDate).slice(0, 10));
      if (!m) return withheld("due-date");
      const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
      const t = new Date(Date.UTC(y, mo - 1, d));
      if (
        y < 2000 || y > 2099 || t.getUTCFullYear() !== y || t.getUTCMonth() !== mo - 1 || t.getUTCDate() !== d
      ) return withheld("due-date");
      due6 = m[1].slice(2) + m[2] + m[3];
    }

    // The versions differ in field order: version 5 puts the RF check digits
    // where version 4 has the reserved 000, and pads after them.
    const built = isRf
      ? "5" + iban.slice(2) + amount8 + ref.slice(2, 4) + ref.slice(4).padStart(21, "0") + due6
      : "4" + iban.slice(2) + amount8 + "000" + v.padStart(20, "0") + due6;
    if (built.length !== 54) return withheld("reference");

    if (erp !== "" && erp !== built) return withheld("mismatch");
    return { outcome: "code", reason: "", iban: printedIban, reference: printedRef, code: built };
  },
});
