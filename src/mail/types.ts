export type MailMetadata = {
  messageRef: string;
  subject: string;
  from: string;
  receivedDateTime: string;
  bodyPreview?: string;
  matchedTerms?: string[];
  bodyReturned: false;
};

export type ReadMailResult = {
  messageRef: string;
  subject: string;
  from: string;
  to: string[];
  cc: string[];
  receivedDateTime: string;
  bodyText: string;
  attachments: Array<{
    name: string;
    contentType: string;
    size: number;
  }>;
  contentWarning: string;
};
