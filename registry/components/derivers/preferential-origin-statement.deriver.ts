import { deriver } from "docxcelerate/template";

/**
 * Decides whether a preferential origin statement can be made, and if so what
 * it says. Installed by `dxcl add preferential-origin-statement`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. The
 * wording, the threshold and every helper are declared inside `run`.
 *
 * Wording sources, to be checked against the text in force before relying on it:
 * EU-UK TCA, Annex ORIG-4; EU-Japan EPA, Annex 3-D; CETA, Protocol on rules of
 * origin, Annex 2 (the bracket label is the likeliest to differ); EU GSP,
 * Commission Implementing Regulation (EU) 2015/2447, Annex 22-07; UK DCTS, the
 * HMRC DCTS origin reference document.
 */
export default deriver({
  name: "preferentialOriginStatement",
  run: (origin: unknown, consignment: unknown) => {
    // deno-lint-ignore no-explicit-any
    const o = origin as any;
    // deno-lint-ignore no-explicit-any
    const c = consignment as any;

    if (typeof o !== "object" || o === null) throw new Error("origin is required");
    const agreements = ["eu-uk-tca", "eu-gsp", "ceta", "eu-japan", "uk-dcts"];
    if (!agreements.includes(o.agreement)) {
      throw new Error("origin.agreement must be one of " + agreements.join(", "));
    }
    const agreement: string = o.agreement;

    const originCountry = typeof o.originCountry === "string" ? o.originCountry.trim() : "";
    const countryOk = agreement === "eu-uk-tca"
      ? originCountry === "UK" || originCountry === "EU"
      : agreement === "ceta" || agreement === "eu-japan"
      ? originCountry === "EU"
      : originCountry !== "" && originCountry !== "EU" && originCountry !== "UK";
    if (!countryOk) {
      throw new Error("origin.originCountry is not allowed for " + agreement + ": " + JSON.stringify(o.originCountry));
    }

    const blank = (v: unknown) => typeof v !== "string" || v.trim() === "";
    if (blank(o.place)) throw new Error("origin.place is required");
    if (blank(o.exporterName)) throw new Error("origin.exporterName is required");
    if (typeof o.date !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(o.date)) {
      throw new Error("origin.date must be an ISO 8601 date");
    }
    const date = o.date.slice(0, 10);

    const currency = c?.currency;
    if (typeof currency !== "string" || !/^[A-Z]{3}$/.test(currency)) {
      throw new Error("consignment.currency must be three upper-case letters");
    }
    if (currency !== "EUR" && !(typeof o.eurRate === "number" && Number.isFinite(o.eurRate) && o.eurRate > 0)) {
      throw new Error("origin.eurRate must be a number above zero when consignment.currency is not EUR");
    }

    const given: unknown = c.lines;
    if (given !== undefined && given !== null && !Array.isArray(given)) {
      throw new Error("consignment.lines must be an array");
    }
    // deno-lint-ignore no-explicit-any
    const lines: any[] = (given as any[] | undefined | null) ?? [];
    const seen = new Set<string>();
    lines.forEach((l, i) => {
      const at = "consignment.lines[" + i + "]";
      if (blank(l?.ref)) throw new Error(at + ".ref is required");
      const key = l.ref.trim();
      if (seen.has(key)) throw new Error(at + ".ref " + JSON.stringify(key) + " is used by another line");
      seen.add(key);
      if (!Number.isInteger(l.valueMinor) || l.valueMinor < 0) {
        throw new Error(at + ".valueMinor must be a non-negative integer");
      }
    });

    const day = (value: unknown, field: string): string => {
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(value)) {
        throw new Error(field + " must be an ISO 8601 date");
      }
      return value.slice(0, 10);
    };
    let from = "";
    let to = "";
    if (o.period !== undefined && o.period !== null) {
      from = day(o.period.from, "origin.period.from");
      to = day(o.period.to, "origin.period.to");
      if (to < from) throw new Error("origin.period.to is before origin.period.from");
    }

    // UTC, so the date does not shift by a day on the engine.
    const formatter = new Intl.DateTimeFormat("en-GB", { dateStyle: "long", timeZone: "UTC" });
    const fmt = (d: string) => formatter.format(new Date(d + "T00:00:00Z"));

    const refOf = (l: { ref: string }) => l.ref.trim();
    const originating = lines.filter((l) => l.originating === true);
    const others = lines.filter((l) => l.originating !== true);

    const title = agreement === "ceta" ? "Origin declaration" : "Statement on origin";
    const exporterName: string = o.exporterName.trim();
    const placeAndDate = o.place.trim() + ", " + fmt(date);

    const originatingMinor: number = originating.reduce((sum, l) => sum + l.valueMinor, 0);
    const digits = new Intl.NumberFormat("en", { style: "currency", currency }).resolvedOptions()
      .maximumFractionDigits as number;
    const rate = currency === "EUR" ? 1 : (o.eurRate as number);
    // Rounded once, here, and nowhere else.
    const originatingEurCents = Math.round(originatingMinor * rate * 100 / 10 ** digits);

    const result = {
      outcome: "none" as "none" | "statement" | "withheld",
      title: "",
      period: null as string | null,
      declaration: "",
      criteriaLine: null as string | null,
      exceptions: null as string | null,
      placeAndDate: "",
      exporterName: "",
      warning: "",
      originatingMinor,
      originatingEurCents,
    };
    if (originating.length === 0) return result;

    const withheld = (warning: string) => ({
      ...result,
      outcome: "withheld" as const,
      title: title + ": not made",
      placeAndDate,
      exporterName,
      warning: "The " + title + " cannot be made: " + warning,
    });

    // 1. A period over 12 months. The last allowed day is one day short of a
    // year on; Date.UTC rolls over month ends, so 2024-02-29 ends 2025-02-28.
    if (from !== "") {
      const last = Date.UTC(+from.slice(0, 4) + 1, +from.slice(5, 7) - 1, +from.slice(8, 10) - 1);
      if (Date.UTC(+to.slice(0, 4), +to.slice(5, 7) - 1, +to.slice(8, 10)) > last) {
        return withheld(
          "the period from " + from + " to " + to + " is longer than the 12 months a statement for multiple " +
            "shipments may cover. Shorten the period and reissue this document.",
        );
      }
    }

    // 2. Origin criteria, where the agreement uses them.
    const agreementName: Record<string, string> = {
      "eu-uk-tca": "the EU-UK Trade and Cooperation Agreement",
      "eu-japan": "the EU-Japan Economic Partnership Agreement",
      "eu-gsp": "the EU Generalised System of Preferences",
      "uk-dcts": "the UK Developing Countries Trading Scheme",
    };
    const letters = agreement === "eu-uk-tca" || agreement === "eu-japan";
    const codes = letters ? "A to E" : "P or W";
    const groups = new Map<string, string[]>();
    if (agreement !== "ceta") {
      for (const l of originating) {
        const raw: unknown = l.criterion;
        const text = typeof raw === "string" ? raw : "";
        let code = "";
        if (letters) {
          const stripped = text.toUpperCase().replace(/[\s,]/g, "");
          if (/^[A-E]+$/.test(stripped)) code = [...new Set(stripped)].sort().join(", ");
        } else {
          const p = text.trim().toUpperCase();
          if (p === "P") code = "P";
          if (p === "W") {
            const heading = typeof l.hsHeading === "string" ? l.hsHeading.trim() : "";
            if (!/^\d{4}$/.test(heading)) {
              return withheld(
                "line " + refOf(l) + " has origin criterion W but no four-digit HS heading. " +
                  "Add the heading and reissue this document.",
              );
            }
            code = "W " + heading;
          }
        }
        if (code === "") {
          return withheld(
            text.trim() === ""
              ? "line " + refOf(l) + " is claimed as originating but has no origin criterion. Add one of " +
                codes + " and reissue this document."
              : "line " + refOf(l) + " has the origin criterion " + JSON.stringify(text) + ", which " +
                agreementName[agreement] + " does not use (" + codes + "). Correct it and reissue this document.",
          );
        }
        groups.set(code, [...(groups.get(code) ?? []), refOf(l)]);
      }
    }

    // 3. The exporter reference. The EUR 6,000 threshold is tested on the
    // originating lines only. It is the same under the EU-UK TCA, EU-Japan,
    // CETA and EU GSP; UK DCTS is the one to check against the DCTS origin
    // reference document.
    const threshold = 600000;
    const reference = typeof o.exporterReference === "string" ? o.exporterReference.trim() : "";
    if (originatingEurCents > threshold && reference === "") {
      const kind = agreement === "eu-uk-tca" && originCountry === "UK" ? "EORI number" : "REX number";
      const eur = new Intl.NumberFormat("en-GB", { style: "currency", currency: "EUR" }).format(
        originatingEurCents / 100,
      );
      return withheld(
        "the originating products in this consignment come to " + eur + ", more than €6,000, so the " +
          "statement must carry the exporter's " + kind + ". Add it and reissue this document.",
      );
    }

    const plural = (refs: string[]) => (refs.length === 1 ? "line " : "lines ") + refs.join(", ");
    const entries = [...groups.entries()];
    const criteria = entries.length === 1
      ? entries[0][0]
      : entries.map(([code, refs]) => code + " (" + plural(refs) + ")").join("; ");

    const label = agreement === "ceta"
      ? "customs or competent governmental authorisation No"
      : agreement === "eu-gsp" || agreement === "uk-dcts"
      ? "Number of Registered Exporter"
      : "Exporter Reference No";
    const refClause = reference ? " (" + label + " " + reference + ")" : "";
    const claim = "declares that, except where otherwise clearly indicated, these products are of " +
      originCountry + " preferential origin";
    const declaration = agreement === "eu-gsp" || agreement === "uk-dcts"
      ? "The exporter " + exporterName + refClause + " of the products covered by this document " +
        claim + " according to rules of origin of the " +
        (agreement === "eu-gsp"
          ? "Generalised System of Preferences of the European Union"
          : "Developing Countries Trading Scheme of the United Kingdom") +
        " and that the origin criterion met is " + criteria + "."
      : "The exporter of the products covered by this document" + refClause + " " + claim + ".";

    return {
      ...result,
      outcome: "statement" as const,
      title,
      period: from === "" ? null : "Period: from " + fmt(from) + " to " + fmt(to),
      declaration,
      criteriaLine: letters ? "Origin criteria: " + criteria : null,
      exceptions: others.length === 0
        ? null
        : "Not of " + originCountry + " preferential origin: " + plural(others.map(refOf)) + ".",
      placeAndDate,
      exporterName,
    };
  },
});
