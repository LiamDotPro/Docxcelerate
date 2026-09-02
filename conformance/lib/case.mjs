/**
 * What a conformance case is, and how it states what it expects.
 *
 * A case is one small document exercising one Word feature, plus assertions
 * over three readings of it: the packed XML, the preview's layout, and Word's
 * own. `defineCase` does nothing but validate the shape and hand it back — the
 * value is in the vocabulary, which is what keeps every case saying the same
 * kind of thing about a different feature.
 *
 * @module
 */

/** The statuses an assertion, a tier or a whole case can hold. */
export const PASS = "PASS";
export const FAIL = "FAIL";
export const KNOWN = "KNOWN";
export const STALE = "STALE";
export const SKIPPED = "SKIPPED";

/** A millimetre, in points and in CSS pixels at 96dpi. */
export const PT_PER_MM = 72 / 25.4;
export const PX_PER_MM = 96 / 25.4;

/**
 * A case, checked for the mistakes that would otherwise show up as a probe
 * crash three steps later.
 *
 * @param {object} spec The case.
 * @returns {object} The same case.
 */
export function defineCase(spec) {
  const required = ["id", "title", "claim", "document"];

  for (const key of required) {
    if (spec[key] === undefined) {
      throw new Error(`case is missing "${key}": ${JSON.stringify(spec.id ?? "(no id)")}`);
    }
  }

  if (!["supported", "partial", "unsupported"].includes(spec.claim)) {
    throw new Error(`${spec.id}: claim must be supported | partial | unsupported`);
  }

  // A `partial` case has to say which tiers are expected to be red, or its
  // status is a shrug rather than a statement.
  if (spec.claim === "partial" && !Array.isArray(spec.knownRed)) {
    throw new Error(`${spec.id}: a partial claim must list knownRed tiers, e.g. ["preview"]`);
  }

  return { knownRed: [], regions: [], ...spec };
}

/**
 * The assertion object handed to every tier.
 *
 * Each call records one objective. Nothing throws: a case that cannot find its
 * paragraph records a miss and carries on, because the first board is expected
 * to be mostly red and a stack trace measures nothing.
 */
export function recorder(tier, unit = "raw") {
  const results = [];
  let n = 0;

  const push = (ok, message, measured, expected, note) => {
    n += 1;
    results.push({
      id: `${tier}.${n}`,
      tier,
      ok,
      message,
      measured: display(measured),
      expected: display(expected),
      ...(note === undefined ? {} : { note }),
    });
    return ok;
  };

  return {
    results,

    /** Deep equality, for a fact that is either right or wrong. */
    equal(measured, expected, message) {
      return push(same(measured, expected), message, measured, expected);
    },

    /** Inequality, for "this must not be what it used to be". */
    notEqual(measured, forbidden, message) {
      return push(!same(measured, forbidden), message, measured, `not ${display(forbidden)}`);
    },

    /**
     * A number inside a tolerance, written as `"1mm"`, `"2pt"`, `"3px"` or a
     * bare number.
     *
     * The tolerance is converted into the tier's own unit, so a case writes
     * `"1mm"` once and every tier means one millimetre of paper by it.
     */
    within(measured, expected, tolerance, message) {
      const slack = toUnits(tolerance, unit);
      const ok = typeof measured === "number" && typeof expected === "number" &&
        Math.abs(measured - expected) <= slack;

      return push(
        ok,
        message,
        measured,
        `${display(expected)} ±${tolerance}`,
        ok || typeof measured !== "number" || typeof expected !== "number"
          ? undefined
          : `off by ${round(Math.abs(measured - expected))}${unit === "raw" ? "" : unit}`,
      );
    },

    /** Strictly greater, for "this row is taller than that one". */
    greater(measured, floor, message) {
      const ok = typeof measured === "number" && measured > floor;
      return push(ok, message, measured, `> ${display(floor)}`);
    },

    /** Strictly less. */
    less(measured, ceiling, message) {
      const ok = typeof measured === "number" && measured < ceiling;
      return push(ok, message, measured, `< ${display(ceiling)}`);
    },

    /** Substring containment, for XML and for text. */
    includes(haystack, needle, message) {
      const ok = typeof haystack === "string" && haystack.includes(needle);
      return push(ok, message, ok ? needle : excerpt(haystack), `contains ${display(needle)}`);
    },

    /** The negative of the above — how "no border was drawn" is stated. */
    excludes(haystack, needle, message) {
      const ok = typeof haystack === "string" && !haystack.includes(needle);
      return push(ok, message, ok ? "absent" : needle, `without ${display(needle)}`);
    },
  };
}

/** Whether two measured facts are the same fact. */
function same(a, b) {
  if (typeof a === "number" && typeof b === "number") {
    return Math.abs(a - b) < 1e-9;
  }
  if (Array.isArray(a) && Array.isArray(b)) {
    return a.length === b.length && a.every((item, index) => same(item, b[index]));
  }
  if (a !== null && b !== null && typeof a === "object" && typeof b === "object") {
    const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
    return [...keys].every((key) => same(a[key], b[key]));
  }
  return a === b;
}

function round(value) {
  return Math.round(value * 100) / 100;
}

/** A measured value, as short as it can be and still be read. */
function display(value) {
  if (value === undefined) return "undefined";
  if (value === null) return "null";
  if (typeof value === "number") return String(round(value));
  if (typeof value === "string") return value.length > 90 ? `${value.slice(0, 87)}…` : value;
  return excerpt(JSON.stringify(value));
}

function excerpt(value) {
  const text = String(value ?? "");
  return text.length > 90 ? `${text.slice(0, 87)}…` : text;
}

/**
 * A tolerance, converted into the unit the tier measures in.
 *
 * A case writes `"1mm"` and means one millimetre of paper wherever it lands:
 * 3.78px on screen, 2.83pt in Word. Exported because the probes want the same
 * arithmetic, and two copies of a unit conversion is one too many.
 */
export function toUnits(tolerance, unit) {
  if (typeof tolerance === "number") return tolerance;
  const match = /^([\d.]+)\s*(mm|pt|px)$/.exec(String(tolerance));
  if (match === null) return Number.parseFloat(tolerance) || 0;

  const value = Number.parseFloat(match[1]);
  const from = match[2];

  const mm = from === "mm" ? value : from === "pt" ? value / PT_PER_MM : value / PX_PER_MM;

  if (unit === "mm") return mm;
  if (unit === "pt") return mm * PT_PER_MM;
  if (unit === "px") return mm * PX_PER_MM;
  return value;
}
