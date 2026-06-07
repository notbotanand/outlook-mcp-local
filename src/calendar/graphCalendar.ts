import { type GraphClient } from "../graph/client.js";
import {
  type CalendarBlockInput,
  type CalendarEventMetadata,
  type CreatedCalendarBlock
} from "./types.js";

type GraphCalendarViewResponse = {
  value: GraphEvent[];
};

type GraphEvent = {
  id: string;
  subject?: string;
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
  location?: { displayName?: string };
  organizer?: { emailAddress?: { name?: string; address?: string } };
  bodyPreview?: string;
};

type GraphCreatedEvent = {
  id: string;
  subject?: string;
  start?: { dateTime?: string };
  end?: { dateTime?: string };
};

export class GraphCalendarService {
  private readonly refs = new Map<string, string>();

  constructor(private readonly graphClient: GraphClient) {}

  async getCalendarView(input: {
    startDateTime: string;
    endDateTime: string;
    maxResults?: number;
  }): Promise<{ results: CalendarEventMetadata[] }> {
    const maxResults = clampMax(input.maxResults, 50);
    const response = await this.graphClient.get<GraphCalendarViewResponse>("/me/calendarView", {
      startDateTime: input.startDateTime,
      endDateTime: input.endDateTime,
      "$select": "id,subject,start,end,location,organizer,bodyPreview",
      "$orderby": "start/dateTime",
      "$top": maxResults
    });

    return {
      results: response.value.slice(0, maxResults).map((event) => this.toMetadata(event))
    };
  }

  async searchCalendar(input: {
    query: string;
    startDateTime: string;
    endDateTime: string;
    maxResults?: number;
  }): Promise<{ results: CalendarEventMetadata[] }> {
    const maxResults = clampMax(input.maxResults, 25);
    const terms = queryTerms(input.query);
    const view = await this.getCalendarView({
      startDateTime: input.startDateTime,
      endDateTime: input.endDateTime,
      maxResults: Math.max(maxResults, 50)
    });

    return {
      results: view.results
        .map((event) => ({ event, matchedTerms: matchedTerms(event, terms) }))
        .filter((candidate) => candidate.matchedTerms.length > 0)
        .slice(0, maxResults)
        .map((candidate) => ({
          ...candidate.event,
          matchedTerms: candidate.matchedTerms
        }))
    };
  }

  async createCalendarBlock(input: CalendarBlockInput): Promise<CreatedCalendarBlock> {
    const timeZone = input.timeZone ?? "UTC";
    const event = await this.graphClient.post<GraphCreatedEvent>("/me/events", {
      subject: input.subject,
      body: {
        contentType: "text",
        content: input.body
      },
      start: {
        dateTime: input.startDateTime,
        timeZone
      },
      end: {
        dateTime: input.endDateTime,
        timeZone
      },
      location: input.location ? { displayName: input.location } : undefined,
      attendees: [],
      isOnlineMeeting: false
    });
    const eventRef = `event_${this.refs.size + 1}`;
    this.refs.set(eventRef, event.id);

    return {
      eventRef,
      subject: event.subject ?? input.subject,
      startDateTime: event.start?.dateTime ?? input.startDateTime,
      endDateTime: event.end?.dateTime ?? input.endDateTime,
      created: true
    };
  }

  private toMetadata(event: GraphEvent): CalendarEventMetadata {
    const eventRef = `event_${this.refs.size + 1}`;
    this.refs.set(eventRef, event.id);

    return {
      eventRef,
      subject: event.subject ?? "",
      startDateTime: event.start?.dateTime ?? "",
      endDateTime: event.end?.dateTime ?? "",
      location: emptyToUndefined(event.location?.displayName),
      organizer: emptyToUndefined(formatOrganizer(event.organizer)),
      bodyPreview: emptyToUndefined(event.bodyPreview),
      bodyReturned: false
    };
  }
}

function clampMax(value: number | undefined, fallback: number): number {
  if (value === undefined || !Number.isFinite(value)) {
    return fallback;
  }

  return Math.max(1, Math.min(100, Math.trunc(value)));
}

function queryTerms(query: string): string[] {
  return query
    .toLowerCase()
    .split(/\s+/)
    .map((term) => term.trim())
    .filter((term) => term.length > 0);
}

function matchedTerms(event: CalendarEventMetadata, terms: string[]): string[] {
  const haystack = [
    event.subject,
    event.location,
    event.organizer,
    event.bodyPreview
  ]
    .filter((value): value is string => value !== undefined)
    .join(" ")
    .toLowerCase();

  return terms.filter((term) => haystack.includes(term));
}

function formatOrganizer(organizer: GraphEvent["organizer"]): string {
  const email = organizer?.emailAddress;
  if (email?.name && email.address) {
    return `${email.name} <${email.address}>`;
  }

  return email?.address ?? email?.name ?? "";
}

function emptyToUndefined(value: string | undefined): string | undefined {
  return value === undefined || value.length === 0 ? undefined : value;
}
