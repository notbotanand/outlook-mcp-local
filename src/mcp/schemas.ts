export type JsonSchema = {
  type: "object";
  properties: Record<string, unknown>;
  required?: string[];
  additionalProperties: false;
};

const isoDateDescription = "ISO 8601 date or datetime string.";
const isoDateTimeDescription = "ISO 8601 datetime string.";

export const outlookSearchMailInputSchema = {
  type: "object",
  properties: {
    query: { type: "string", minLength: 1 },
    startDate: { type: "string", description: isoDateDescription },
    endDate: { type: "string", description: isoDateDescription },
    maxResults: { type: "number", minimum: 1, maximum: 100, default: 10 }
  },
  required: ["query"],
  additionalProperties: false
} satisfies JsonSchema;

export const outlookListMailInputSchema = {
  type: "object",
  properties: {
    startDate: { type: "string", description: isoDateDescription },
    endDate: { type: "string", description: isoDateDescription },
    folder: { type: "string" },
    maxResults: { type: "number", minimum: 1, maximum: 100, default: 25 }
  },
  required: ["startDate", "endDate"],
  additionalProperties: false
} satisfies JsonSchema;

export const outlookReadMailInputSchema = {
  type: "object",
  properties: {
    messageRef: { type: "string", pattern: "^msg_[A-Za-z0-9_-]+$" }
  },
  required: ["messageRef"],
  additionalProperties: false
} satisfies JsonSchema;

export const outlookGetCalendarViewInputSchema = {
  type: "object",
  properties: {
    startDateTime: { type: "string", description: isoDateTimeDescription },
    endDateTime: { type: "string", description: isoDateTimeDescription },
    maxResults: { type: "number", minimum: 1, maximum: 100, default: 50 }
  },
  required: ["startDateTime", "endDateTime"],
  additionalProperties: false
} satisfies JsonSchema;

export const outlookSearchCalendarInputSchema = {
  type: "object",
  properties: {
    query: { type: "string", minLength: 1 },
    startDateTime: { type: "string", description: isoDateTimeDescription },
    endDateTime: { type: "string", description: isoDateTimeDescription },
    maxResults: { type: "number", minimum: 1, maximum: 100, default: 25 }
  },
  required: ["query", "startDateTime", "endDateTime"],
  additionalProperties: false
} satisfies JsonSchema;

export const outlookCreateCalendarBlockInputSchema = {
  type: "object",
  properties: {
    subject: { type: "string", minLength: 1 },
    body: { type: "string" },
    startDateTime: { type: "string", description: isoDateTimeDescription },
    endDateTime: { type: "string", description: isoDateTimeDescription },
    timeZone: { type: "string" },
    location: { type: "string" }
  },
  required: ["subject", "body", "startDateTime", "endDateTime"],
  additionalProperties: false
} satisfies JsonSchema;

export const authStatusInputSchema = {
  type: "object",
  properties: {},
  additionalProperties: false
} satisfies JsonSchema;

export const authLoginInputSchema = {
  type: "object",
  properties: {
    flow: { type: "string", enum: ["pkce", "device"], default: "pkce" }
  },
  additionalProperties: false
} satisfies JsonSchema;

export const authCompleteInputSchema = {
  type: "object",
  properties: {
    authSessionId: { type: "string", minLength: 1 }
  },
  required: ["authSessionId"],
  additionalProperties: false
} satisfies JsonSchema;

export const authLogoutInputSchema = {
  type: "object",
  properties: {},
  additionalProperties: false
} satisfies JsonSchema;

export const storageStatusInputSchema = {
  type: "object",
  properties: {},
  additionalProperties: false
} satisfies JsonSchema;
