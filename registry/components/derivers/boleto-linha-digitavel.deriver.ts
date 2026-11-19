// deno-lint-ignore-file no-explicit-any -- the inputs are request data nobody has checked.
import { deriver } from "docxcelerate/template";

/**
 * The typeable line of a boleto, built from the invoice and, when the bank
 * returned one, checked against it digit for digit. Installed by
 * `dxcl add boleto-linha-digitavel`.
 *
 * Self-contained on purpose: it travels to the engine as source text, so a
 * module-level helper or constant would be a ReferenceError there. Every
 * helper is declared inside `run`.
 *
 * The first rule that applies decides: a malformed due date or bank line
 * throws, a line that cannot be built is withheld with its reason, and
 * otherwise the line is built and, if given, cross-checked.
 */
export default deriver({
  name: "boletoLinhaDigitavel",
  run: (boleto: unknown, invoice: unknown) => {
    const b = boleto as any;
    const inv = invoice as any;

    // Weights run from the rightmost digit leftwards. Modulus 11 is the
    // general check digit, over the barcode order; 0, 10 and 11 become 1.
    const mod11 = (digits: string): number => {
      let sum = 0;
      for (let i = 0; i < digits.length; i++) {
        sum += Number(digits[digits.length - 1 - i]) * (2 + (i % 8));
      }
      const dv = 11 - (sum % 11);
      return dv > 9 ? 1 : dv;
    };
    // Modulus 10 is the field check digit, over the line order.
    const mod10 = (digits: string): number => {
      let sum = 0;
      for (let i = 0; i < digits.length; i++) {
        const product = Number(digits[digits.length - 1 - i]) * (i % 2 === 0 ? 2 : 1);
        sum += product > 9 ? product - 9 : product;
      }
      return (10 - (sum % 10)) % 10;
    };
    const money = (cents: number): string =>
      new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" }).format(cents / 100);
    const brDate = (iso: string): string => iso.slice(8, 10) + "/" + iso.slice(5, 7) + "/" + iso.slice(0, 4);
    const epoch = Date.UTC(1997, 9, 7);
    const dayMs = 86400000;
    const withheldOutput = (reason: string) => ({
      line: null,
      barcode: null,
      dueDate: null,
      amount: null,
      withheld: "Linha digitável indisponível: " + reason + ". Solicite o boleto ao beneficiário.",
    });

    // 1. The due date, when there is one, must be a real date.
    const rawDue = inv?.dueDate;
    const hasDue = rawDue !== undefined && rawDue !== null && rawDue !== "";
    let dueIso = "";
    let days = 0;
    if (hasDue) {
      dueIso = String(rawDue).slice(0, 10);
      const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dueIso);
      const t = m ? Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : NaN;
      if (!m || new Date(t).toISOString().slice(0, 10) !== dueIso) {
        throw new Error("invoice.dueDate must be an ISO 8601 date");
      }
      days = Math.round((t - epoch) / dayMs);
    }

    // 2. The bank's line, when given, must be 47 digits.
    const given = b?.line === undefined || b?.line === null ? "" : String(b.line);
    let L = "";
    if (given !== "") {
      L = given.replace(/[ .]/g, "");
      if (!/^\d{47}$/.test(L)) throw new Error("boleto.line must be 47 digits");
    }

    // 3. Withhold, first reason wins.
    const bank = typeof b?.bankCode === "string" ? b.bankCode.trim() : "";
    const free = typeof b?.freeField === "string" ? b.freeField.trim() : "";
    const total = inv?.total;
    if (inv?.currency !== "BRL") return withheldOutput("a fatura não está em reais (BRL)");
    if (!/^\d{3}$/.test(bank)) return withheldOutput("o código do banco não tem 3 dígitos");
    if (!/^\d{25}$/.test(free)) return withheldOutput("o campo livre não tem 25 dígitos");
    if (hasDue && days < 1000) {
      return withheldOutput(
        "o vencimento é anterior a 03/07/2000, a primeira data que o fator de vencimento representa",
      );
    }
    if (typeof total !== "number" || !Number.isFinite(total)) {
      return withheldOutput("o total da fatura não é um valor");
    }
    // Scaling through the string avoids 1.005 * 100 = 100.49999999999999.
    const cents = Math.round(Number(String(total) + "e2"));
    if (!(cents >= 1)) return withheldOutput("não há valor a pagar");
    if (cents > 9999999999) {
      return withheldOutput("o valor passa de R$ 99.999.999,99, o máximo que a linha comporta");
    }

    // 4. Build. The factor restarts at 1000 every 9000 days, from 2025-02-22.
    const factor = hasDue ? String(((days - 1000) % 9000) + 1000).padStart(4, "0") : "0000";
    const amountField = String(cents).padStart(10, "0");
    const dv = mod11(bank + "9" + factor + amountField + free);
    const barcode = bank + "9" + dv + factor + amountField + free;
    const f1 = bank + "9" + free.slice(0, 5);
    const f2 = free.slice(5, 15);
    const f3 = free.slice(15, 25);
    const digits = f1 + mod10(f1) + f2 + mod10(f2) + f3 + mod10(f3) + dv + factor + amountField;

    // 5. Cross-check the bank's line. It collects what its line says.
    if (L !== "") {
      for (const [n, from, to] of [[1, 0, 9], [2, 10, 20], [3, 21, 31]] as const) {
        if (mod10(L.slice(from, to)) !== Number(L[to])) {
          throw new Error("boleto.line fails the check digit of field " + n + "; it is mistyped or corrupt");
        }
      }
      const lineFree = L.slice(4, 9) + L.slice(10, 20) + L.slice(21, 31);
      const lineBarcode = L.slice(0, 4) + L[32] + L.slice(33) + lineFree;
      if (mod11(lineBarcode.slice(0, 4) + lineBarcode.slice(5)) !== Number(L[32])) {
        throw new Error("boleto.line fails its general check digit");
      }
      if (L.slice(0, 3) !== bank) {
        throw new Error("boleto.line is for bank " + L.slice(0, 3) + "; the invoice is for bank " + bank);
      }
      if (L[3] !== "9") throw new Error("boleto.line is not in reais (currency code " + L[3] + ")");
      if (lineFree !== free) throw new Error("boleto.line belongs to a different boleto (free field differs)");
      const lineFactor = L.slice(33, 37);
      if (lineFactor !== factor) {
        const describe = (f: string, target: number | null): string => {
          if (f === "0000") return "has no due date";
          const k = target === null ? 1 : Math.max(0, Math.round((target - Number(f)) / 9000));
          return "is due " + brDate(new Date(epoch + (Number(f) + 9000 * k) * dayMs).toISOString().slice(0, 10));
        };
        throw new Error(
          "boleto.line " + describe(lineFactor, hasDue ? days : null) + "; the invoice " +
            describe(factor, hasDue ? days : null),
        );
      }
      const lineAmount = L.slice(37);
      if (lineAmount !== amountField) {
        const describe = (a: string): string =>
          Number(a) === 0 ? "no amount (a variable-amount boleto)" : money(Number(a));
        throw new Error("boleto.line is for " + describe(lineAmount) + "; the invoice is for " + describe(amountField));
      }
      if (L !== digits) throw new Error("boleto.line does not match the invoice");
    }

    return {
      line: digits.slice(0, 5) + "." + digits.slice(5, 10) + " " + digits.slice(10, 15) + "." +
        digits.slice(15, 21) + " " + digits.slice(21, 26) + "." + digits.slice(26, 32) + " " + digits[32] + " " +
        digits.slice(33),
      barcode,
      dueDate: hasDue ? brDate(dueIso) : "Contra apresentação",
      amount: money(cents),
      withheld: null,
    };
  },
});
