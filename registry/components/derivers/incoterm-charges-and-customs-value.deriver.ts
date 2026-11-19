// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Splits a shipment's cost legs by who pays them under the Incoterms 2020 rule,
 * totals the invoice, and states the customs value on the destination's basis.
 * Installed by `dxcl add incoterm-charges-and-customs-value`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. The
 * matrix, labels, formatters and rounding all live inside `run`.
 *
 * Reading the matrix:
 *
 * - Packing is the seller's under every rule (A8).
 * - FCA at the seller's own premises has no origin-inland leg: the author
 *   leaves it out.
 * - Under CPT and CIP to an inland place, carriage to that place is all
 *   main-carriage.
 * - Insurance under the D-rules is the seller's because the seller carries the
 *   risk. It is not obliged, which is why a missing leg there means
 *   "uninsured", not "incomplete".
 * - Destination handling that the seller's freight contract already includes
 *   belongs in main-carriage.
 */
export default deriver({
  name: "incotermChargesAndCustomsValue",
  run: (shipment: unknown) => {
    const s = shipment as any;

    const rules = ["EXW", "FCA", "FAS", "FOB", "CFR", "CIF", "CPT", "CIP", "DAP", "DPU", "DDP"];
    // One column per rule, in the order of `rules`: S seller, B buyer.
    const matrix: Record<string, string> = {
      "export-packing": "SSSSSSSSSSS",
      "origin-inland": "BSSSSSSSSSS",
      "export-clearance": "BSSSSSSSSSS",
      "origin-terminal": "BBBSSSSSSSS",
      "main-carriage": "BBBBSSSSSSS",
      "insurance": "BBBBBSBSSSS",
      "destination-terminal": "BBBBBBBBSSS",
      "onward-delivery": "BBBBBBBBSSS",
      "import-duty": "BBBBBBBBBBS",
      "import-tax": "BBBBBBBBBBS",
    };
    const labels: Record<string, string> = {
      "export-packing": "Export packing",
      "origin-inland": "Inland transport to the place of export",
      "export-clearance": "Export clearance",
      "origin-terminal": "Terminal handling at origin",
      "main-carriage": "Main carriage (international freight)",
      "insurance": "Cargo insurance",
      "destination-terminal": "Terminal handling at destination",
      "onward-delivery": "Delivery from the place of arrival",
      "import-duty": "Import duty",
      "import-tax": "Import VAT and taxes",
    };
    const kinds = Object.keys(matrix);
    const modes = ["sea", "sea-container", "air", "road", "rail", "multimodal"];
    const modeWords: Record<string, string> = {
      "sea-container": "containerised",
      air: "air",
      road: "road",
      rail: "rail",
      multimodal: "multimodal",
    };

    const rule: string = s?.incoterm;
    const ruleIndex = rules.indexOf(rule);
    if (ruleIndex < 0) throw new Error("shipment.incoterm must be one of the eleven Incoterms 2020 rules");

    const currency = s.currency;
    let money: Intl.NumberFormat;
    try {
      // Intl accepts any three letters, so a code it does not list is refused by hand.
      if (typeof currency !== "string" || !(Intl as any).supportedValuesOf("currency").includes(currency)) {
        throw new Error("unknown");
      }
      money = new Intl.NumberFormat("en-GB", { style: "currency", currency });
    } catch {
      throw new Error("shipment.currency must be an ISO 4217 code");
    }
    const exponent = money.resolvedOptions().maximumFractionDigits as number;
    const fmt = (minor: number) => money.format(minor / 10 ** exponent);

    const minorOf = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isInteger(value) || value < 0) {
        throw new Error(field + " must be a non-negative integer of minor units");
      }
      return value;
    };
    const goodsMinor = minorOf(s.goodsMinor, "shipment.goodsMinor");

    if (s.legs !== undefined && !Array.isArray(s.legs)) throw new Error("shipment.legs must be an array");
    const mode: string | undefined = s.mode;
    if (mode !== undefined && !modes.includes(mode)) {
      throw new Error("shipment.mode must be one of " + modes.join(", "));
    }

    const dest = s.destination;
    const basis = dest?.valuationBasis;
    if (basis !== "cif-border" && basis !== "fob-export") {
      throw new Error('shipment.destination.valuationBasis must be "cif-border" or "fob-export"');
    }
    const pctGiven = dest.airFreightDutiablePercent;
    if (pctGiven !== undefined && (typeof pctGiven !== "number" || !Number.isInteger(pctGiven) || pctGiven < 0 || pctGiven > 100)) {
      throw new Error("shipment.destination.airFreightDutiablePercent must be an integer from 0 to 100");
    }

    // Validate every leg, then index by kind.
    const byKind: Record<string, { amount: number; inPrice: boolean; seller: boolean }> = {};
    ((s.legs ?? []) as any[]).forEach((leg, i) => {
      const f = "shipment.legs[" + i + "]";
      if (!kinds.includes(leg?.kind)) throw new Error(f + ".kind is not a known leg");
      if (byKind[leg.kind]) throw new Error(f + ".kind " + leg.kind + " appears twice; add the amounts together");
      const amount = minorOf(leg.amountMinor, f + ".amountMinor");
      const seller = matrix[leg.kind][ruleIndex] === "S";
      const inPrice = leg.inPrice === true;
      if (inPrice && !seller) throw new Error(f + ".inPrice is true, but " + rule + " puts " + leg.kind + " on the buyer");
      byKind[leg.kind] = { amount, inPrice, seller };
    });

    let included = 0;
    let charged = 0;
    for (const k of kinds) {
      const l = byKind[k];
      if (!l || !l.seller) continue;
      if (l.inPrice) included += l.amount;
      else charged += l.amount;
    }
    if (included > goodsMinor) {
      throw new Error("the legs included in the price come to more than shipment.goodsMinor");
    }
    const totalMinor = goodsMinor + charged;

    // FOB: take the seller's post-export legs out, put the buyer's pre-export legs in.
    const afterExport = ["main-carriage", "insurance", "destination-terminal", "onward-delivery", "import-duty", "import-tax"];
    const beforeExport = ["export-packing", "origin-inland", "export-clearance", "origin-terminal"];
    let fobMinor = totalMinor;
    for (const k of kinds) {
      const l = byKind[k];
      if (!l) continue;
      if (l.seller && afterExport.includes(k)) fobMinor -= l.amount;
      if (!l.seller && beforeExport.includes(k)) fobMinor += l.amount;
    }

    const mainAmount = byKind["main-carriage"]?.amount ?? 0;
    const insuranceAmount = byKind["insurance"]?.amount ?? 0;
    // Round half up on exact integers, once, on the apportioned freight.
    const pct = pctGiven ?? 100;
    const dutiableFreight = basis === "cif-border" && mode === "air"
      ? Math.floor((2 * mainAmount * pct + 100) / 200)
      : mainAmount;
    const cifComputable = !(basis === "cif-border" && !byKind["main-carriage"]);
    const cifMinor = cifComputable ? fobMinor + dutiableFreight + insuranceAmount : null;

    let reason = "";
    if (basis === "cif-border") {
      if (!byKind["main-carriage"]) reason = "the cost of main carriage to the place of entry was not given";
    } else if (matrix["main-carriage"][ruleIndex] === "S" && !byKind["main-carriage"]) {
      reason = "the price includes main carriage, and it has to be itemised to be taken out";
    } else if ((rule === "CIF" || rule === "CIP") && !byKind["insurance"]) {
      reason = "the price includes insurance, and it has to be itemised to be taken out";
    }
    const customsMinor = reason ? null : basis === "cif-border" ? cifMinor : fobMinor;

    const customsLine = reason
      ? "Customs value: not stated, because " + reason + "."
      : basis === "cif-border"
      ? "Customs value, CIF to the place of entry: " + fmt(customsMinor as number) +
        ". Transport and handling after entry, and import duty and taxes, are shown separately above and are not part of it."
      : "Customs value, FOB at the place of export: " + fmt(customsMinor as number) + ".";

    let customsNote = "";
    if (basis === "fob-export") {
      customsNote =
        "International freight and insurance are itemised above by name and amount, and are not part of this value (19 CFR 141.86(a)(8)).";
    } else if (mode === "air") {
      customsNote = pctGiven === undefined
        ? "Air freight counted in full. Supply the Annex 23-01 percentage for the country of departure to apportion it."
        : "Air freight counted at " + pctGiven + "% under Annex 23-01 of Implementing Regulation (EU) 2015/2447.";
    }

    const place = typeof s.namedPlace === "string" ? s.namedPlace.trim() : "";
    const term = rule + (place ? " " + place : "") + " Incoterms® 2020";

    const seaOnly = ["FAS", "FOB", "CFR", "CIF"].includes(rule);
    const replacement: Record<string, string> = { FAS: "FCA", FOB: "FCA", CFR: "CPT", CIF: "CIP" };
    const warnSeaOnly = seaOnly && mode !== undefined && mode !== "sea"
      ? rule + " is a rule for sea and inland waterway transport only. For " + modeWords[mode] +
        " freight the rule meant is " + replacement[rule] + ": " + rule +
        " leaves the goods at the seller's risk until they are on board, long after the carrier has taken them."
      : "";
    const warnNamedPlace = place
      ? ""
      : "No named place is given for " + rule + ", so the invoice does not say where costs and risk pass to the buyer.";
    const warnDdpVat = rule === "DDP" && byKind["import-tax"]
      ? "Under DDP the seller pays the import VAT, which a seller not registered in the destination usually cannot recover. Consider DAP, or agree that import VAT is for the buyer's account."
      : "";
    const warnImpliedInsurance = (rule === "CFR" || rule === "CPT") && byKind["insurance"]
      ? rule + " does not oblige the seller to insure the goods. This invoice shows insurance, which implies they are insured. If the seller insured them, the rule is " +
        (rule === "CFR" ? "CIF" : "CIP") + "."
      : "";

    const legs: Record<string, { label: string; treatment: string; amount: string }> = {};
    for (const k of kinds) {
      const l = byKind[k];
      legs[k] = l
        ? {
          label: labels[k],
          treatment: !l.seller ? "For buyer's account" : l.inPrice ? "Included in price" : "Charged",
          amount: fmt(l.amount),
        }
        : { label: "", treatment: "", amount: "" };
    }

    return {
      term,
      goods: fmt(goodsMinor),
      legs,
      total: fmt(totalMinor),
      customsLine,
      customsNote,
      warnSeaOnly,
      warnNamedPlace,
      warnDdpVat,
      warnImpliedInsurance,
      totalMinor,
      fobMinor,
      cifMinor,
      customsMinor,
    };
  },
});
