// ─── Shared TypeScript Types ──────────────────────────────────────────────────

export type AssistantMode = "site" | "web";

// ── Chat ──────────────────────────────────────────────────────────────────────

export type ChatEntryType = "text" | "email_draft" | "error" | "thinking";

export interface EmailDraftData {
  to: string[];
  cc: string[];
  subject: string;
  plainText: string;
  htmlBody: string;
  sent: boolean;
  cancelled: boolean;
  loading: boolean;
}

export interface ChatEntry {
  id: string;
  role: "user" | "assistant";
  text: string;
  type: ChatEntryType;
  emailData?: EmailDraftData;
  timestamp: number;
}

// ── Graph ─────────────────────────────────────────────────────────────────────

export interface SearchResult {
  title: string;
  snippet: string;
  webUrl: string;
  isExternal?: boolean;
  lastModified?: string;
}

export interface PageContext {
  title: string;
  textContent: string;
  url: string;
}

export interface UserProfile {
  displayName: string;
  mail: string;
  jobTitle?: string;
  department?: string;
}

export interface ManagerProfile {
  displayName: string;
  mail: string;
}

export interface EmailDraft {
  to: string[];
  cc: string[];
  subject: string;
  htmlBody: string;
}
