import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { GraphCalendarService } from "../../src/calendar/graphCalendar.js";
import { type GraphClient } from "../../src/graph/client.js";

describe("GraphCalendarService", () => {
  it("returns calendar view metadata only", async () => {
    const calls: Array<{ path: string; query?: Record<string, string | number | undefined> }> = [];
    const client: GraphClient = {
      async get<T>(path: string, query?: Record<string, string | number | undefined>): Promise<T> {
        calls.push({ path, query });
        return {
          value: [
            {
              id: "graph-event-1",
              subject: "Contoso interview prep",
              start: { dateTime: "2026-06-06T10:00:00" },
              end: { dateTime: "2026-06-06T11:00:00" },
              location: { displayName: "Home office" },
              organizer: { emailAddress: { name: "Alex Johnson", address: "alex.johnson@example.com" } },
              bodyPreview: "Prepare notes for Contoso conversation."
            }
          ]
        } as T;
      },
      async post<T>(): Promise<T> {
        throw new Error("not used");
      }
    };

    const service = new GraphCalendarService(client);
    const result = await service.getCalendarView({
      startDateTime: "2026-06-06T00:00:00-07:00",
      endDateTime: "2026-06-07T00:00:00-07:00"
    });

    assert.equal(calls[0].path, "/me/calendarView");
    assert.equal(calls[0].query?.startDateTime, "2026-06-06T00:00:00-07:00");
    assert.equal(result.results[0].eventRef, "event_1");
    assert.equal(result.results[0].bodyReturned, false);
    assert.equal("bodyText" in result.results[0], false);
  });

  it("searches calendar locally over metadata fields", async () => {
    const client: GraphClient = {
      async get<T>(): Promise<T> {
        return {
          value: [
            {
              id: "graph-event-1",
              subject: "Contoso interview prep",
              start: { dateTime: "2026-06-06T10:00:00" },
              end: { dateTime: "2026-06-06T11:00:00" },
              bodyPreview: "Prepare notes for Contoso conversation."
            },
            {
              id: "graph-event-2",
              subject: "Fabrikam onboarding",
              start: { dateTime: "2026-06-06T14:00:00" },
              end: { dateTime: "2026-06-06T15:00:00" },
              bodyPreview: "Synthetic onboarding design review."
            }
          ]
        } as T;
      },
      async post<T>(): Promise<T> {
        throw new Error("not used");
      }
    };

    const service = new GraphCalendarService(client);
    const result = await service.searchCalendar({
      query: "Contoso notes",
      startDateTime: "2026-06-06T00:00:00-07:00",
      endDateTime: "2026-06-07T00:00:00-07:00"
    });

    assert.equal(result.results.length, 1);
    assert.equal(result.results[0].subject, "Contoso interview prep");
    assert.deepEqual(result.results[0].matchedTerms, ["contoso", "notes"]);
    assert.equal(result.results[0].bodyReturned, false);
  });

  it("creates personal calendar blocks without attendees or recurrence", async () => {
    let postedPath = "";
    let postedBody: unknown;
    const client: GraphClient = {
      async get<T>(): Promise<T> {
        throw new Error("not used");
      },
      async post<T>(path: string, body: unknown): Promise<T> {
        postedPath = path;
        postedBody = body;
        return {
          id: "graph-event-3",
          subject: "Focus block",
          start: { dateTime: "2026-06-06T16:00:00" },
          end: { dateTime: "2026-06-06T17:00:00" }
        } as T;
      }
    };

    const service = new GraphCalendarService(client);
    const result = await service.createCalendarBlock({
      subject: "Focus block",
      body: "Synthetic prep notes",
      startDateTime: "2026-06-06T16:00:00",
      endDateTime: "2026-06-06T17:00:00",
      timeZone: "America/Los_Angeles",
      location: "Home office"
    });

    assert.equal(postedPath, "/me/events");
    assert.equal(result.created, true);
    assert.equal(result.eventRef, "event_1");
    assert.deepEqual(postedBody, {
      subject: "Focus block",
      body: {
        contentType: "text",
        content: "Synthetic prep notes"
      },
      start: {
        dateTime: "2026-06-06T16:00:00",
        timeZone: "America/Los_Angeles"
      },
      end: {
        dateTime: "2026-06-06T17:00:00",
        timeZone: "America/Los_Angeles"
      },
      location: { displayName: "Home office" },
      attendees: [],
      isOnlineMeeting: false
    });
  });
});
