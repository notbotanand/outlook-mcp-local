import assert from "node:assert/strict";
import { describe, it } from "node:test";
import calendarFixtures from "../fixtures/synthetic/calendar.json" with { type: "json" };
import mailFixtures from "../fixtures/synthetic/mail.json" with { type: "json" };
import { callTool, getEnabledTools } from "../../src/mcp/tools.js";

describe("MCP tools", () => {
  it("uses synthetic fixtures only", () => {
    assert.equal(mailFixtures[0].from, "recruiting@example.com");
    assert.equal(calendarFixtures[0].organizer, "alex.johnson@example.com");
  });

  it("hides calendar block creation in read mode", () => {
    const names = getEnabledTools("read").map((tool) => tool.name);

    assert.equal(names.includes("outlook_create_calendar_block"), false);
    assert.ok(names.includes("outlook_search_mail"));
    assert.ok(names.includes("outlook_get_calendar_view"));
  });

  it("exposes calendar block creation in calendar-write mode", () => {
    const names = getEnabledTools("calendar-write").map((tool) => tool.name);

    assert.ok(names.includes("outlook_create_calendar_block"));
  });

  it("requires authentication for mail search without a mail service", async () => {
    const result = await callTool({ mode: "read" }, "outlook_search_mail", { query: "Contoso" });

    assert.equal(result.isError, true);
  });

  it("requires authentication for calendar view without a calendar service", async () => {
    const result = await callTool(
      { mode: "read" },
      "outlook_get_calendar_view",
      {
        startDateTime: "2026-06-06T00:00:00-07:00",
        endDateTime: "2026-06-07T00:00:00-07:00"
      }
    );

    assert.equal(result.isError, true);
  });

  it("keeps mail search metadata-only through injected mail service", async () => {
    const result = await callTool(
      {
        mode: "read",
        mailService: {
          async searchMail() {
            return {
              results: [
                {
                  messageRef: "msg_1",
                  subject: "Contoso",
                  from: "recruiting@example.com",
                  receivedDateTime: "2026-06-05T09:00:00Z",
                  bodyPreview: "Synthetic preview",
                  bodyReturned: false
                }
              ]
            };
          },
          async listMail() {
            return { results: [] };
          },
          async readMail() {
            throw new Error("not used");
          }
        }
      },
      "outlook_search_mail",
      { query: "Contoso" }
    );
    const payload = JSON.parse(result.content[0].text) as {
      results: Array<{ bodyReturned: boolean; bodyText?: string }>;
    };

    assert.equal(result.isError, undefined);
    assert.equal(payload.results[0].bodyReturned, false);
    assert.equal(payload.results[0].bodyText, undefined);
  });

  it("blocks calendar block tool calls in read mode", async () => {
    const result = await callTool(
      { mode: "read" },
      "outlook_create_calendar_block",
      {
        subject: "Focus block",
        body: "Synthetic note",
        startDateTime: "2026-06-06T10:00:00-07:00",
        endDateTime: "2026-06-06T11:00:00-07:00"
      }
    );

    assert.equal(result.isError, true);
  });

  it("keeps calendar search metadata-only through injected calendar service", async () => {
    const result = await callTool(
      {
        mode: "read",
        calendarService: {
          async getCalendarView() {
            return { results: [] };
          },
          async searchCalendar() {
            return {
              results: [
                {
                  eventRef: "event_1",
                  subject: "Contoso interview prep",
                  startDateTime: "2026-06-06T10:00:00",
                  endDateTime: "2026-06-06T11:00:00",
                  matchedTerms: ["contoso"],
                  bodyReturned: false
                }
              ]
            };
          },
          async createCalendarBlock() {
            throw new Error("not used");
          }
        }
      },
      "outlook_search_calendar",
      {
        query: "Contoso",
        startDateTime: "2026-06-06T00:00:00-07:00",
        endDateTime: "2026-06-07T00:00:00-07:00"
      }
    );
    const payload = JSON.parse(result.content[0].text) as {
      results: Array<{ bodyReturned: boolean; bodyText?: string }>;
    };

    assert.equal(result.isError, undefined);
    assert.equal(payload.results[0].bodyReturned, false);
    assert.equal(payload.results[0].bodyText, undefined);
  });

  it("creates calendar block in calendar-write mode through injected calendar service", async () => {
    const result = await callTool(
      {
        mode: "calendar-write",
        calendarService: {
          async getCalendarView() {
            return { results: [] };
          },
          async searchCalendar() {
            return { results: [] };
          },
          async createCalendarBlock(input) {
            return {
              eventRef: "event_1",
              subject: input.subject,
              startDateTime: input.startDateTime,
              endDateTime: input.endDateTime,
              created: true
            };
          }
        }
      },
      "outlook_create_calendar_block",
      {
        subject: "Focus block",
        body: "Synthetic prep",
        startDateTime: "2026-06-06T16:00:00",
        endDateTime: "2026-06-06T17:00:00"
      }
    );
    const payload = JSON.parse(result.content[0].text) as { created: boolean };

    assert.equal(result.isError, undefined);
    assert.equal(payload.created, true);
  });

  it("exposes auth helper tools in read mode", () => {
    const names = getEnabledTools("read").map((tool) => tool.name);

    assert.ok(names.includes("auth_status"));
    assert.ok(names.includes("auth_login"));
    assert.ok(names.includes("auth_complete"));
    assert.ok(names.includes("auth_logout"));
    assert.ok(names.includes("storage_status"));
  });

  it("reports storage status without values", async () => {
    const result = await callTool(
      {
        mode: "read",
        tokenStore: {
          async read() {
            return undefined;
          },
          async write() {
            return undefined;
          },
          async clear() {
            return undefined;
          },
          describe() {
            return {
              refreshTokenStore: "process memory",
              accessTokenStore: "process memory"
            };
          }
        }
      },
      "storage_status",
      {}
    );
    const payload = JSON.parse(result.content[0].text) as {
      tokenStore: { refreshTokenStore: string };
      serverPersistsMailBodies: boolean;
    };

    assert.equal(result.isError, undefined);
    assert.equal(payload.tokenStore.refreshTokenStore, "process memory");
    assert.equal(payload.serverPersistsMailBodies, false);
  });
});
