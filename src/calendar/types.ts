export type CalendarEventMetadata = {
  eventRef: string;
  subject: string;
  startDateTime: string;
  endDateTime: string;
  location?: string;
  organizer?: string;
  bodyPreview?: string;
  matchedTerms?: string[];
  bodyReturned: false;
};

export type CalendarBlockInput = {
  subject: string;
  body: string;
  startDateTime: string;
  endDateTime: string;
  timeZone?: string;
  location?: string;
};

export type CreatedCalendarBlock = {
  eventRef: string;
  subject: string;
  startDateTime: string;
  endDateTime: string;
  created: true;
};
