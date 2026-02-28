import { graphPost } from "./graphClient";
import type { EmailDraft } from "./types";
import { trackEvent, trackError } from "../utils/telemetry";

// ─── Graph Mail Service ───────────────────────────────────────────────────────
// Sends email in the delegated context of the current user.
// NEVER called without explicit user confirmation (Confirm & Send button).

export async function sendMail(draft: EmailDraft): Promise<void> {
  try {
    await graphPost("/me/sendMail", {
      message: {
        subject: draft.subject,
        body: {
          contentType: "HTML",
          content: draft.htmlBody,
        },
        toRecipients: draft.to.map((address) => ({
          emailAddress: { address },
        })),
        ccRecipients: draft.cc.map((address) => ({
          emailAddress: { address },
        })),
      },
      saveToSentItems: true,
    });

    trackEvent({
      name: "email_sent",
      properties: { toCount: draft.to.length, subject: draft.subject },
    });
  } catch (error) {
    trackError("graphMailService.sendMail", error, { subject: draft.subject });
    throw error;
  }
}
