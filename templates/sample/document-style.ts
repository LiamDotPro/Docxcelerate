import { cleanMinimalDocumentStyle, type DocumentStyle } from "docxcelerate/document";

export const documentStyle: DocumentStyle = {
  ...cleanMinimalDocumentStyle,
  page: {
    ...cleanMinimalDocumentStyle.page,
    margins: {
      topMm: 25.4,
      rightMm: 25.4,
      bottomMm: 25.4,
      leftMm: 25.4,
    },
  },
};
