import { type JsonSchema } from "./schemas.js";

export type ValidationResult =
  | { valid: true }
  | { valid: false; message: string };

type SchemaProperty = {
  type?: string;
  enum?: unknown[];
  minLength?: number;
  minimum?: number;
  maximum?: number;
  pattern?: string;
};

export function validateInputSchema(schema: JsonSchema, value: unknown): ValidationResult {
  if (!isRecord(value)) {
    return { valid: false, message: "Tool arguments must be an object." };
  }

  for (const key of schema.required ?? []) {
    if (value[key] === undefined) {
      return { valid: false, message: `Missing required argument: ${key}` };
    }
  }

  if (schema.additionalProperties === false) {
    for (const key of Object.keys(value)) {
      if (schema.properties[key] === undefined) {
        return { valid: false, message: `Unknown argument: ${key}` };
      }
    }
  }

  for (const [key, rawProperty] of Object.entries(schema.properties)) {
    const candidate = value[key];
    if (candidate === undefined) {
      continue;
    }

    const result = validateProperty(key, candidate, rawProperty as SchemaProperty);
    if (!result.valid) {
      return result;
    }
  }

  return { valid: true };
}

export function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function validateProperty(key: string, value: unknown, property: SchemaProperty): ValidationResult {
  if (property.type === "string") {
    if (typeof value !== "string") {
      return { valid: false, message: `Argument ${key} must be a string.` };
    }

    if (property.minLength !== undefined && value.length < property.minLength) {
      return { valid: false, message: `Argument ${key} is too short.` };
    }

    if (property.pattern !== undefined && !new RegExp(property.pattern).test(value)) {
      return { valid: false, message: `Argument ${key} has an invalid format.` };
    }
  }

  if (property.type === "number") {
    if (typeof value !== "number" || !Number.isFinite(value)) {
      return { valid: false, message: `Argument ${key} must be a number.` };
    }

    if (property.minimum !== undefined && value < property.minimum) {
      return { valid: false, message: `Argument ${key} is below the minimum.` };
    }

    if (property.maximum !== undefined && value > property.maximum) {
      return { valid: false, message: `Argument ${key} is above the maximum.` };
    }
  }

  if (property.enum !== undefined && !property.enum.includes(value)) {
    return { valid: false, message: `Argument ${key} is not an allowed value.` };
  }

  return { valid: true };
}
