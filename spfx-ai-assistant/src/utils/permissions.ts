// ─── Permission & Access Message Helpers ────────────────────────────────────

export const NO_RESULTS_MESSAGE =
  "No relevant content found in this site.";

export const RESTRICTED_ACCESS_MESSAGE =
  "A matching document may exist, but you may not have permission to access it.";

export const EXTERNAL_SOURCE_LABEL = "External Source";

export const GRAPH_ERROR_MESSAGES: Record<number, string> = {
  401: "Authentication failed. Please refresh the page and try again.",
  403: "You do not have permission to access this resource.",
  404: "The requested resource was not found.",
  429: "Too many requests. Please wait a moment and try again.",
  500: "A server error occurred. Please try again later.",
};

export function getGraphErrorMessage(status: number): string {
  return (
    GRAPH_ERROR_MESSAGES[status] ||
    `An unexpected error occurred (HTTP ${status}).`
  );
}

/**
 * Determine whether the error looks like an access/permission denial.
 */
export function isAccessDenied(error: unknown): boolean {
  if (!error) return false;
  const msg = String((error as { message?: string }).message || "").toLowerCase();
  return (
    msg.includes("403") ||
    msg.includes("forbidden") ||
    msg.includes("unauthorized") ||
    msg.includes("access denied")
  );
}

/**
 * Build a user-friendly message for empty search results.
 */
export function buildEmptyResultsMessage(
  query: string,
  accessDenied: boolean
): string {
  if (accessDenied) {
    return `${NO_RESULTS_MESSAGE} ${RESTRICTED_ACCESS_MESSAGE}`;
  }
  return `${NO_RESULTS_MESSAGE} Try rephrasing your query: "${query}".`;
}
