import { readFile } from "node:fs/promises";
import { resolve } from "node:path";

import { config as loadDotenv } from "dotenv";
import { z } from "zod";

loadDotenv({ quiet: true });

const serviceAccountFileSchema = z.object({
  client_email: z.string().min(1),
  private_key: z.string().min(1),
  project_id: z.string().min(1).optional(),
});

interface GoogleSheetsInlineCredentials {
  clientEmail: string;
  privateKey: string;
  projectId?: string;
}

export interface GoogleSheetsSyncConfig extends GoogleSheetsInlineCredentials {
  spreadsheetId: string;
}

function normalizeEnvValue(value: string | undefined): string | undefined {
  const normalized = value?.trim();
  return normalized ? normalized : undefined;
}

function normalizePrivateKey(value: string): string {
  return value.replace(/\\n/g, "\n");
}

function readInlineCredentials(): GoogleSheetsInlineCredentials | null {
  const clientEmail = normalizeEnvValue(process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL);
  const privateKey = normalizeEnvValue(process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY);
  const projectId = normalizeEnvValue(process.env.GOOGLE_SERVICE_ACCOUNT_PROJECT_ID);

  if (!clientEmail && !privateKey && !projectId) {
    return null;
  }

  if (!clientEmail || !privateKey) {
    throw new Error(
      "Incomplete Google Sheets credentials in .env. Set GOOGLE_SERVICE_ACCOUNT_EMAIL and GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY.",
    );
  }

  return {
    clientEmail,
    privateKey: normalizePrivateKey(privateKey),
    projectId,
  };
}

async function readServiceAccountFile(): Promise<GoogleSheetsInlineCredentials | null> {
  const keyPath = normalizeEnvValue(process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH);
  if (!keyPath) {
    return null;
  }

  const filePath = resolve(keyPath);
  const raw = await readFile(filePath, "utf8");
  const parsed = serviceAccountFileSchema.parse(JSON.parse(raw));

  return {
    clientEmail: parsed.client_email,
    privateKey: normalizePrivateKey(parsed.private_key),
    projectId: parsed.project_id,
  };
}

export async function resolveGoogleSheetsSyncConfig(): Promise<GoogleSheetsSyncConfig | null> {
  const spreadsheetId = normalizeEnvValue(process.env.GOOGLE_SHEETS_SPREADSHEET_ID);
  const inlineCredentials = readInlineCredentials();
  const fileCredentials = await readServiceAccountFile();

  if (!spreadsheetId && !inlineCredentials && !fileCredentials) {
    return null;
  }

  if (!spreadsheetId) {
    throw new Error(
      "Incomplete Google Sheets configuration in .env. Set GOOGLE_SHEETS_SPREADSHEET_ID.",
    );
  }

  const credentials = inlineCredentials ?? fileCredentials;
  if (!credentials) {
    throw new Error(
      "Google Sheets credentials are not configured. Set GOOGLE_SERVICE_ACCOUNT_EMAIL and GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY, or GOOGLE_SERVICE_ACCOUNT_KEY_PATH.",
    );
  }

  return {
    spreadsheetId,
    ...credentials,
  };
}
