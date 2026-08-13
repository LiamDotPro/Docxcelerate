export interface OfferData {
  applicantName: string;
  programme: string;
  college: string;
  startDate: string;
  offerRef: string;
  conditions: string[];
  replyBy: string;
  /** Context for the tutor's note — never printed verbatim. */
  portfolioTheme: string;
  interviewer: string;
  tuitionFee: string;
  feeStatus: "Home" | "International";
  scholarship?: { name: string; amount: string };
  signatory: { name: string; title: string };
}
