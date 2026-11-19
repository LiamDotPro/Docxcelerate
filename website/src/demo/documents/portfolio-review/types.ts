/** The firm sending the review. Printed on every page. */
export interface Firm {
  name: string;
  /** The single letter in the mark's box. */
  initial: string;
  /** "Authorised and regulated by …" — the footer's left half. */
  regulation: string;
  address: string;
  phone: string;
  registeredOffice: string;
}

/** Who the review is for. */
export interface Client {
  /** "Ms Adaeze Nkemelu" — the cover and the address block. */
  name: string;
  /** "Ms A. Nkemelu" — the running header, where space is short. */
  shortName: string;
  /** "Ms Nkemelu" — the letter's greeting. */
  salutation: string;
  addressLines: string[];
  portfolio: string;
  account: string;
}

/** The adviser who signs the review. */
export interface Adviser {
  name: string;
  role: string;
  phone: string;
  /** The signature, set in a script face. */
  signature: string;
}

/** The quarter the review covers. */
export interface Period {
  /** "Third quarter 2026". */
  title: string;
  /** "Q3 2026". */
  short: string;
  /** "1 July to 30 September 2026". */
  range: string;
  /** "30 September" — the date every value is taken at. */
  valuedAt: string;
  /** "30 June" — the date the quarter opened. */
  openedAt: string;
  /** "7 October 2026" — when the review was written. */
  preparedOn: string;
  reference: string;
}

/** One figure in the summary grid. */
export interface KeyFigure {
  label: string;
  value: string;
  note: string;
  /** Whether the note is a rise, printed in the accent with an arrow. */
  rise?: boolean;
}

export interface Trade {
  side: "Sell" | "Buy";
  fund: string;
  amount: number;
}

export interface NextStep {
  title: string;
  body: string;
}

/** One row of the returns table. */
export interface ReturnRow {
  label: string;
  values: string[];
}

export interface CalendarYear {
  year: string;
  portfolio: number;
  benchmark: number;
}

/** A contribution to the quarter's return, in percentage points. */
export interface Contribution {
  name: string;
  points: number;
}

/** An asset class against its agreed weight. All in percent. */
export interface AllocationClass {
  name: string;
  now: number;
  target: number;
  /** Half the agreed range: 3 means target ± 3 points. */
  tolerance: number;
  after: number;
}

export interface Share {
  name: string;
  /** What the stacked bar prints inside the segment, if it is wide enough. */
  code?: string;
  percent: number;
}

/** One line of the holdings table. */
export interface Holding {
  name: string;
  kind: string;
  weight: number;
  value: number;
  quarter: number;
  year: number;
  /** The fund manager's note, handed to the engine for "why it moved". */
  managerNote: string;
}

export interface Account {
  name: string;
  value: number;
  note: string;
}

export interface Fact {
  label: string;
  value: string;
}

/** A holding given a panel of its own on the focus pages. */
export interface FocusHolding {
  name: string;
  kind: string;
  value: number;
  weight: number;
  /** Twelve monthly values, any scale: only the shape is drawn. */
  spark: number[];
  facts: Fact[];
  positions: Share[];
  /** What the fund is for, for the engine to explain per client. */
  purpose: string;
}

export interface MonthIncome {
  month: string;
  amount: number;
  /** Whether the month falls in the quarter under review. */
  inQuarter: boolean;
}

export interface Amount {
  label: string;
  amount: number;
}

/** One step of the value bridge from opening to closing value. */
export interface BridgeStep {
  label: string;
  /** Signed change; the first and last steps are totals instead. */
  amount: number;
  total?: boolean;
}

export interface Transaction {
  date: string;
  description: string;
  holding: string;
  amount: number;
}

export interface CostRow {
  charge: string;
  how: string;
  amount: number;
  percent: number;
}

export interface Projection {
  years: string[];
  /** In thousands of pounds. */
  central: number[];
  strong: number[];
  weak: number[];
  goal: number;
}

export interface Allowance {
  name: string;
  used: number;
  limit: number;
}

export interface GlossaryTerm {
  term: string;
  meaning: string;
}

export interface PortfolioReviewData {
  firm: Firm;
  client: Client;
  adviser: Adviser;
  period: Period;
  nextReview: string;
  openingValue: number;
  closingValue: number;
  /** The quarter's return, in pounds and percent, and the benchmark's. */
  quarterReturn: number;
  quarterPercent: number;
  benchmarkPercent: number;
  keyFigures: KeyFigure[];
  drift: string;
  trades: Trade[];
  nextSteps: NextStep[];
  /** Twelve months rebased to 100. */
  months: string[];
  portfolioIndex: number[];
  benchmarkIndex: number[];
  returnColumns: string[];
  returns: ReturnRow[];
  calendarYears: CalendarYear[];
  byClass: Contribution[];
  byHolding: Contribution[];
  allocation: AllocationClass[];
  regions: Share[];
  currencies: Share[];
  holdings: Holding[];
  accounts: Account[];
  focus: FocusHolding[];
  income: MonthIncome[];
  incomeByHolding: Amount[];
  cash: Amount[];
  bridge: BridgeStep[];
  transactions: Transaction[];
  costs: CostRow[];
  yearCosts: Amount[];
  costEffect: { years: string[]; before: number[]; after: number[] };
  projection: Projection;
  allowances: Allowance[];
  disclosures: string[];
  glossary: GlossaryTerm[];
  /** What the client has told us about their plans, for the engine. */
  clientNotes: string;
}
