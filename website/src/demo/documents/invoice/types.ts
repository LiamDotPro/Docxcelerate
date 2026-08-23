/** One billable line of the engagement. */
export interface InvoiceLine {
  /** What the work was. */
  desc: string;
  /** The note printed under it — which sprint, which system. */
  meta: string;
  /** Days, or whatever the rate is per. */
  qty: number;
  /** The rate that quantity is charged at. */
  rate: number;
}

/** Where money is sent, printed on the payment page. */
export interface BankDetails {
  accountName: string;
  sortCode: string;
  accountNumber: string;
  iban: string;
  bic: string;
}

/** The party sending the invoice. */
export interface Sender {
  name: string;
  /** The line under the name — what they do, and where. */
  trade: string;
  addressLines: string[];
  email: string;
  /** Company number and VAT number, printed in the footer. */
  registration: string;
  bank: BankDetails;
}

/** The party being billed. */
export interface BilledTo {
  name: string;
  /** Who the invoice is for the attention of. */
  attn: string;
  addressLines: string[];
}

export interface InvoiceData {
  reference: string;
  issueDate: string;
  dueDate: string;
  poReference: string;
  /** Whether this has been settled. Decides the status the invoice carries. */
  paid: boolean;
  /** Whether the footer credits the tool that wrote this. */
  showCredit: boolean;
  sender: Sender;
  billedTo: BilledTo;
  lines: InvoiceLine[];
  /** As a fraction — 0.2 for the UK's twenty percent. */
  vatRate: number;
  /** Who queries about the invoice go to. */
  deliveryLead: string;
}
