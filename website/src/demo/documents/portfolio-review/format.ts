/**
 * How the review writes numbers: whole pounds, a true minus sign, and a sign
 * on every change so that a gain and a loss never read alike.
 */

const grouped = new Intl.NumberFormat("en-GB", { maximumFractionDigits: 0 });

/** £412,860, or −£2,316. */
export const pounds = (amount: number): string =>
  `${amount < 0 ? "−" : ""}£${grouped.format(Math.abs(amount))}`;

/** +£1,500 or −£86: a movement, always signed. */
export const signedPounds = (amount: number): string =>
  `${amount < 0 ? "−" : "+"}£${grouped.format(Math.abs(amount))}`;

/** +1.59 or −0.07, to a fixed number of places. */
export const signed = (value: number, places = 2): string =>
  `${value < 0 ? "−" : "+"}${Math.abs(value).toFixed(places)}`;

/** 64.2%, to a fixed number of places. */
export const percent = (value: number, places = 1): string => `${value.toFixed(places)}%`;

/** ▲ 5.6% or ▼ 0.6%: a return, with its direction drawn. */
export const arrowed = (value: number, places = 1): string =>
  `${value < 0 ? "▼" : "▲"} ${Math.abs(value).toFixed(places)}%`;
