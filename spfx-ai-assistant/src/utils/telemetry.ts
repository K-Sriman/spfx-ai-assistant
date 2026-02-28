// ─── Prototype Telemetry (console-only) ─────────────────────────────────────
// In production, replace with Azure Application Insights or similar.

const PREFIX = "[AI-Assistant]";

export type TelemetryEventName =
  | "panel_opened"
  | "panel_closed"
  | "mode_switched"
  | "query_submitted"
  | "knowledge_search_complete"
  | "email_intent_detected"
  | "email_draft_shown"
  | "email_sent"
  | "email_cancelled"
  | "web_search_complete"
  | "error";

interface TelemetryEvent {
  name: TelemetryEventName;
  properties?: Record<string, string | number | boolean>;
}

export function trackEvent(event: TelemetryEvent): void {
  const { name, properties } = event;
  const ts = new Date().toISOString();
  // eslint-disable-next-line no-console
  console.log(`${PREFIX} [${ts}] EVENT: ${name}`, properties || "");
}

export function trackError(
  source: string,
  error: unknown,
  extra?: Record<string, string>
): void {
  const message =
    error instanceof Error ? error.message : String(error);
  // eslint-disable-next-line no-console
  console.error(`${PREFIX} [ERROR] ${source}:`, message, extra || "");
}

export function trackInfo(message: string, data?: unknown): void {
  // eslint-disable-next-line no-console
  console.info(`${PREFIX}`, message, data !== undefined ? data : "");
}
