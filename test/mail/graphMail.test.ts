import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { GraphMailService } from "../../src/mail/graphMail.js";
import { type GraphClient } from "../../src/graph/client.js";

describe("GraphMailService", () => {
  it("returns metadata only for list results and reads explicit refs", async () => {
    const calls: string[] = [];
    const client: GraphClient = {
      async get<T>(path: string): Promise<T> {
        calls.push(path);

        if (path === "/me/messages") {
          return {
            value: [
              {
                id: "graph-message-1",
                subject: "Contoso follow-up",
                from: { emailAddress: { name: "Recruiting", address: "recruiting@example.com" } },
                receivedDateTime: "2026-06-05T09:00:00Z",
                bodyPreview: "Synthetic preview"
              }
            ]
          } as T;
        }

        return {
          id: "graph-message-1",
          subject: "Contoso follow-up",
          from: { emailAddress: { address: "recruiting@example.com" } },
          toRecipients: [{ emailAddress: { address: "alex.johnson@example.com" } }],
          ccRecipients: [],
          receivedDateTime: "2026-06-05T09:00:00Z",
          body: { contentType: "html", content: "<p>Synthetic body</p>" },
          attachments: [{ name: "agenda.txt", contentType: "text/plain", size: 12 }]
        } as T;
      },
      async post<T>(): Promise<T> {
        throw new Error("not used");
      }
    };

    const service = new GraphMailService(client);
    const listed = await service.listMail({
      startDate: "2026-06-05T00:00:00Z",
      endDate: "2026-06-06T00:00:00Z"
    });

    assert.equal(listed.results[0].messageRef, "msg_1");
    assert.equal(listed.results[0].bodyReturned, false);
    assert.equal("bodyText" in listed.results[0], false);

    const read = await service.readMail({ messageRef: "msg_1" });
    assert.equal(read.bodyText, "Synthetic body");
    assert.equal(read.attachments[0].name, "agenda.txt");
    assert.deepEqual(calls, ["/me/messages", "/me/messages/graph-message-1"]);
  });
});
