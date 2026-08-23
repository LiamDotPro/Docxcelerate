import type { JsonObject } from "../domain/types.ts";

/**
 * The component registry: prebuilt nodes, described once.
 *
 * The entries here are metadata. The code itself lives under `registry/` at the
 * root of the package, as ordinary `.tsx` files that import `docxcelerate` the
 * way a project would — which means they are typechecked against the real
 * framework rather than kept as strings, and `dxcl add` installs them by
 * copying rather than by generating.
 *
 * Copying is the point. An installed component is your file: it has no version,
 * it is not upgraded behind you, and editing it is the expected next step
 * rather than a fork. What the registry gives you is the first draft and the
 * decisions already made in it.
 *
 * @module
 */

/** Where in a document a component usually sits. */
export type ComponentCategory =
  | "Opening"
  | "Body"
  | "Closing"
  | "Legal"
  | "Generated";

/** One field an installed component expects to find on the project's data. */
export interface RegistryDataField {
  /** The path, as it is read — `sender.name`, not `name`. */
  path: string;
  /** The TypeScript type the component reads it as. */
  type: string;
  /** What it is for, and what happens when it is missing. */
  summary: string;
}

/** One file an install copies. */
export interface RegistryFile {
  /** Path under `registry/`, inside the package. */
  source: string;
  /** Where it lands, relative to the document project. */
  target: string;
}

/** A prebuilt component, as the CLI and the site read it. */
export interface RegistryComponent {
  /** Discriminator, so themes and components can share one lookup. */
  kind: "component";
  /** The id: the URL slug, and what `dxcl add` is given. */
  id: string;
  /** The component's name, as it is printed. */
  title: string;
  /** One line, for cards and lists. */
  summary: string;
  /** A paragraph: what it does, and the decision it has already made for you. */
  detail: string;
  /** Where in a document it usually sits. */
  category: ComponentCategory;
  /** Free-form labels a reader might filter on. */
  tags: string[];
  /** The component names the files export, in the order they should be used. */
  exports: string[];
  /** What it reads, so a project knows what to add to its data type. */
  dataFields: RegistryDataField[];
  /** The files copied in. */
  files: RegistryFile[];
  /** Data the previews on the site are built against, and a project can crib. */
  previewData: JsonObject;
  /** Other registry ids installed alongside this one. */
  requires: string[];
  /** The theme it was drawn against, where it was drawn against one. */
  themeHint?: string;
}

/**
 * One file per component, named after the component, under `nodes/`.
 *
 * Flat rather than a directory each: a document project keeps every node in one
 * place, and an installed component is meant to stop being distinguishable from
 * one you wrote as soon as you have edited it.
 */
function node(id: string): RegistryFile[] {
  return [{ source: `components/${id}/${id}.node.tsx`, target: `nodes/${id}.node.tsx` }];
}

export const COMPONENTS: RegistryComponent[] = [
  {
    kind: "component",
    id: "letterhead",
    title: "Letterhead",
    summary: "Who sent this, from where, and when.",
    detail:
      "The block at the top of the page. The address is one paragraph with " +
      "the lines joined rather than one node per line, so a sender with two " +
      "address lines and one with five both come out as a block instead of a " +
      "ragged run of nodes.",
    category: "Opening",
    tags: ["letter", "header", "static"],
    exports: ["Letterhead"],
    dataFields: [
      {
        path: "sender.name",
        type: "string",
        summary: "The organisation, as it should be printed.",
      },
      {
        path: "sender.addressLines",
        type: "string[]",
        summary: "Street, town, postcode — one entry per line, joined for print.",
      },
      {
        path: "sentOn",
        type: "string | number | Date",
        summary: "When the document was sent. Formatted for the build locale.",
      },
    ],
    files: node("letterhead"),
    previewData: {
      sender: {
        name: "Riverside Leisure Centre",
        addressLines: ["14 Mill Lane", "Bristol", "BS1 4TG"],
      },
      sentOn: "2026-09-01",
    },
    requires: [],
  },
  {
    kind: "component",
    id: "recipient-block",
    title: "Recipient block",
    summary: "The address, and a greeting that survives a missing name.",
    detail:
      "The salutation is the half worth having. A document generated in bulk " +
      "meets recipients whose names it does not have — a joint tenancy, a " +
      "company, a record where the field was never filled in — so the " +
      "greeting branches, and the fallback is formal rather than clever.",
    category: "Opening",
    tags: ["letter", "address", "branching"],
    exports: ["RecipientBlock"],
    dataFields: [
      {
        path: "recipient.name",
        type: "string | undefined",
        summary: "The addressee. Absent is handled: the greeting falls back.",
      },
      {
        path: "recipient.addressLines",
        type: "string[]",
        summary: "Street, town, postcode — one entry per line.",
      },
      {
        path: "recipient.formalName",
        type: "string | undefined",
        summary: "Used in the greeting where a first name would be too familiar.",
      },
    ],
    files: node("recipient-block"),
    previewData: {
      recipient: {
        name: "Adaeze Nkemelu",
        formalName: "Ms Nkemelu",
        addressLines: ["Flat 6", "22 Colston Street", "Bristol", "BS1 5AE"],
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "payment-summary",
    title: "Payment summary",
    summary: "What is owed, by when — and what to say when nothing is.",
    detail:
      "Three outcomes, three ids: in credit, clear, or owing. Branching on " +
      "the amount rather than papering over it with one sentence that reads " +
      "oddly at zero is the point — somebody who owes nothing should not be " +
      "given a payment deadline.",
    category: "Body",
    tags: ["money", "branching", "section"],
    exports: ["PaymentSummary"],
    dataFields: [
      {
        path: "account.reference",
        type: "string",
        summary: "Your reference, printed so a caller can quote it.",
      },
      {
        path: "account.balanceDue",
        type: "number",
        summary: "What is owed. Negative means in credit; zero means clear.",
      },
      {
        path: "account.dueBy",
        type: "string | number | Date | undefined",
        summary: "When it is due. Printed only where something is owed.",
      },
      {
        path: "account.currency",
        type: "string | undefined",
        summary: "ISO 4217 code. Defaults to GBP.",
      },
    ],
    files: node("payment-summary"),
    previewData: {
      account: {
        reference: "RIV-88214",
        balanceDue: 128.42,
        dueBy: "2026-09-30",
        currency: "GBP",
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "next-steps",
    title: "Next steps",
    summary: "A generated paragraph with all four prompts already fenced off.",
    detail:
      "The node most likely to invent something, so all four prompts are set: " +
      "the info prompt hands the engine the facts it may use, and the " +
      "negative prompt closes off the two failures that matter in a document " +
      "somebody acts on — an invented deadline and an invented way to " +
      "contact you.",
    category: "Generated",
    tags: ["dynamic", "prompts", "ai"],
    exports: ["NextSteps"],
    dataFields: [
      {
        path: "actions",
        type: "string[]",
        summary: "What the reader has to do, in your words. Given to the engine as fact.",
      },
      {
        path: "contact",
        type: "string",
        summary: "How to reach you. Given to the engine rather than printed.",
      },
    ],
    files: node("next-steps"),
    previewData: {
      actions: [
        "check the balance shown above",
        "pay online or by bank transfer",
        "tell us if anything looks wrong",
      ],
      contact: "hello@riversideleisure.example on 0117 496 0000",
    },
    requires: [],
  },
  {
    kind: "component",
    id: "signature-block",
    title: "Signature block",
    summary: "A closing, a signature image, and who signed it.",
    detail:
      "The image is optional and the name is not, which is the right way " +
      "round: a letter signed by nobody is a letter nobody owns, whereas a " +
      "missing image is a rendering detail. Without one the block closes with " +
      "the typed name.",
    category: "Closing",
    tags: ["letter", "image", "closing"],
    exports: ["SignatureBlock"],
    dataFields: [
      {
        path: "signatory.name",
        type: "string",
        summary: "The person signing, as they should be printed.",
      },
      {
        path: "signatory.role",
        type: "string",
        summary: "Their role, printed beneath the name.",
      },
      {
        path: "signatory.signatureImage",
        type: "string | undefined",
        summary: "Path to a signature image, relative to the document project.",
      },
      {
        path: "signatory.closing",
        type: "string | undefined",
        summary: 'How the letter closes. Defaults to "Yours sincerely".',
      },
    ],
    files: node("signature-block"),
    previewData: {
      signatory: {
        name: "Tomas Lindqvist",
        role: "Centre Manager",
        signatureImage: "assets/signature-lindqvist.png",
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "terms-notice",
    title: "Terms notice",
    summary: "Numbered small print, one paragraph and one stable id per clause.",
    detail:
      "Boilerplate is the part of a document nobody rereads and everybody " +
      "copies, so it is worth one node that owns it. Clauses arrive as data, " +
      "each becomes its own paragraph, and each id is stable — which is what " +
      "lets next year's terms be diffed against this year's clause by clause.",
    category: "Legal",
    tags: ["legal", "boilerplate", "section"],
    exports: ["TermsNotice"],
    dataFields: [
      {
        path: "terms.heading",
        type: "string | undefined",
        summary: 'The heading above the clauses. Defaults to "Terms".',
      },
      {
        path: "terms.clauses",
        type: "string[]",
        summary: "One entry per clause, printed in order and numbered.",
      },
      {
        path: "terms.version",
        type: "string | undefined",
        summary: "Version or date, printed last so a reader can cite it.",
      },
    ],
    files: node("terms-notice"),
    previewData: {
      terms: {
        heading: "Terms of membership",
        clauses: [
          "Membership runs for twelve months from the renewal date shown above.",
          "Prices are held for the term and reviewed once a year.",
          "You may cancel within fourteen days of renewal for a full refund.",
        ],
        version: "2026.1",
      },
    },
    requires: [],
    themeHint: "legal-serif",
  },
];

/** Kept in this order everywhere the catalog is listed. */
export const COMPONENT_CATEGORIES: ComponentCategory[] = [
  "Opening",
  "Body",
  "Generated",
  "Closing",
  "Legal",
];
