export interface RepairsData {
  residentName: string;
  address: string;
  jobRef: string;
  trade: "plumber" | "electrician" | "joiner";
  visitDate: string;
  visitWindow: string;
  accessNotes?: string;
}
