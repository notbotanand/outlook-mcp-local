import { type GraphClient } from "../graph/client.js";
import { emailContentWarning } from "./constants.js";
import { htmlToText } from "./sanitizer.js";
import { type MailMetadata, type ReadMailResult } from "./types.js";

type GraphMailResponse = {
  value: GraphMessage[];
};

type GraphMessage = {
  id: string;
  subject?: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  toRecipients?: GraphRecipient[];
  ccRecipients?: GraphRecipient[];
  receivedDateTime?: string;
  bodyPreview?: string;
  body?: { contentType?: string; content?: string };
  attachments?: GraphAttachment[];
};

type GraphRecipient = {
  emailAddress?: {
    name?: string;
    address?: string;
  };
};

type GraphAttachment = {
  name?: string;
  contentType?: string;
  size?: number;
};

export class GraphMailService {
  private readonly refs = new Map<string, string>();

  constructor(private readonly graphClient: GraphClient) {}

  async searchMail(input: {
    query: string;
    startDate?: string;
    endDate?: string;
    maxResults?: number;
  }): Promise<{ results: MailMetadata[] }> {
    const maxResults = clampMax(input.maxResults, 10);
    const response = await this.graphClient.get<GraphMailResponse>("/me/messages", {
      "$select": "id,subject,from,receivedDateTime,bodyPreview",
      "$search": quoteSearch(input.query),
      "$top": maxResults
    });
    const terms = input.query.split(/\s+/).filter(Boolean);

    return {
      results: response.value
        .filter((message) => withinRange(message.receivedDateTime, input.startDate, input.endDate))
        .slice(0, maxResults)
        .map((message) => this.toMetadata(message, terms))
    };
  }

  async listMail(input: {
    startDate: string;
    endDate: string;
    folder?: string;
    maxResults?: number;
  }): Promise<{ results: MailMetadata[] }> {
    const maxResults = clampMax(input.maxResults, 25);
    const path = input.folder ? `/me/mailFolders/${encodeURIComponent(input.folder)}/messages` : "/me/messages";
    const response = await this.graphClient.get<GraphMailResponse>(path, {
      "$select": "id,subject,from,receivedDateTime,bodyPreview",
      "$filter": `receivedDateTime ge ${input.startDate} and receivedDateTime le ${input.endDate}`,
      "$orderby": "receivedDateTime desc",
      "$top": maxResults
    });

    return {
      results: response.value.slice(0, maxResults).map((message) => this.toMetadata(message))
    };
  }

  async readMail(input: { messageRef: string }): Promise<ReadMailResult> {
    const graphId = this.refs.get(input.messageRef);
    if (graphId === undefined) {
      throw new Error("Unknown messageRef. Search or list mail before reading a message.");
    }

    const message = await this.graphClient.get<GraphMessage>(`/me/messages/${encodeURIComponent(graphId)}`, {
      "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body",
      "$expand": "attachments($select=name,contentType,size)"
    });

    return {
      messageRef: input.messageRef,
      subject: message.subject ?? "",
      from: formatRecipient(message.from),
      to: (message.toRecipients ?? []).map(formatRecipient),
      cc: (message.ccRecipients ?? []).map(formatRecipient),
      receivedDateTime: message.receivedDateTime ?? "",
      bodyText: bodyToText(message.body),
      attachments: (message.attachments ?? []).map((attachment) => ({
        name: attachment.name ?? "",
        contentType: attachment.contentType ?? "",
        size: attachment.size ?? 0
      })),
      contentWarning: emailContentWarning
    };
  }

  private toMetadata(message: GraphMessage, matchedTerms?: string[]): MailMetadata {
    const messageRef = `msg_${this.refs.size + 1}`;
    if (message.id !== undefined) {
      this.refs.set(messageRef, message.id);
    }

    return {
      messageRef,
      subject: message.subject ?? "",
      from: formatRecipient(message.from),
      receivedDateTime: message.receivedDateTime ?? "",
      bodyPreview: message.bodyPreview,
      matchedTerms,
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

function quoteSearch(query: string): string {
  return `"${query.replaceAll('"', '\\"')}"`;
}

function withinRange(value: string | undefined, start?: string, end?: string): boolean {
  if (value === undefined) {
    return true;
  }

  return (start === undefined || value >= start) && (end === undefined || value <= end);
}

function formatRecipient(recipient: GraphRecipient | GraphMessage["from"] | undefined): string {
  const email = recipient?.emailAddress;
  if (email?.name && email.address) {
    return `${email.name} <${email.address}>`;
  }

  return email?.address ?? email?.name ?? "";
}

function bodyToText(body: GraphMessage["body"]): string {
  const content = body?.content ?? "";
  if (body?.contentType?.toLowerCase() === "html") {
    return htmlToText(content);
  }

  return content;
}
