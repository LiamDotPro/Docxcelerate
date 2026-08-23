import type { InvoiceData } from "./types.ts";

export const previewData: InvoiceData = {
  reference: "INV-2026-0142",
  issueDate: "2026-08-21",
  dueDate: "2026-09-04",
  poReference: "PO-BHA-2214",
  paid: false,
  showCredit: true,
  sender: {
    name: "Fernhill Systems Ltd",
    trade: "Software consultancy · Manchester",
    addressLines: ["Unit 9, Carding Mill", "Manchester M4 5JW"],
    email: "accounts@fernhill.systems",
    registration: "Registered in England & Wales No. 09184472 · VAT GB 312 4477 08",
    bank: {
      accountName: "Fernhill Systems Ltd",
      sortCode: "04-00-72",
      accountNumber: "18732209",
      iban: "GB29 FRNH 0400 7218 7322 09",
      bic: "FRNHGB2L",
    },
  },
  billedTo: {
    name: "Brackenfield Housing Association",
    attn: "Attn Maya Oyelaran, Finance",
    addressLines: ["4 Millrace Court", "Leeds LS2 7QF"],
  },
  lines: [
    { desc: "Discovery and scoping workshop", meta: "Tenant portal programme", qty: 2, rate: 760 },
    { desc: "API development — Sprint 14", meta: "Repairs booking endpoints", qty: 8, rate: 760 },
    { desc: "Tenant portal front-end build", meta: "Account and statements views", qty: 6, rate: 760 },
    {
      desc: "Document automation",
      meta: "Docxcelerate templates: arrears letters, statements",
      qty: 4,
      rate: 760,
    },
    { desc: "CI and release automation", meta: "GitHub Actions, staged deploys", qty: 1.5, rate: 760 },
    { desc: "Accessibility audit and fixes", meta: "WCAG 2.2 AA across the portal", qty: 2, rate: 680 },
    { desc: "Production support retainer", meta: "August 2026", qty: 1, rate: 950 },
  ],
  vatRate: 0.2,
  deliveryLead: "Priya Raman",
};
