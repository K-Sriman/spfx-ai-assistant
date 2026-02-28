import type { PageContext } from "./types";
import { trackError } from "../utils/telemetry";

// ─── SharePoint Page Service ─────────────────────────────────────────────────

/**
 * Extract the current page's content from the DOM.
 * Prefers the SharePoint canvas content region, falls back to <main> or <body>.
 */
export function extractPageContent(): PageContext {
  try {
    const title = document.title.replace(" - SharePoint", "").trim();
    const url = location.href;

    // Try SharePoint modern canvas first
    const selectors = [
      "#spPageCanvasContent",
      "[data-automation-id='pageContent']",
      ".CanvasComponent",
      "main[role='main']",
      "main",
      "#contentRow",
    ];

    let textContent = "";
    for (const selector of selectors) {
      const el = document.querySelector(selector);
      if (el && el.textContent && el.textContent.trim().length > 100) {
        textContent = el.textContent.trim();
        break;
      }
    }

    // Final fallback: body minus nav/header/footer
    if (!textContent) {
      const body = document.body.cloneNode(true) as HTMLElement;
      // Remove navigation and chrome elements
      ["nav", "header", "footer", ".od-TopBar", "#sp-appBar", "#spLeftNav"].forEach((sel) => {
        body.querySelectorAll(sel).forEach((n) => n.parentNode?.removeChild(n));
      });
      textContent = body.textContent?.trim() || "";
    }

    // Cap at 5000 chars to stay within context limits
    return {
      title,
      textContent: textContent.slice(0, 5000),
      url,
    };
  } catch (error) {
    trackError("spPageService.extractPageContent", error);
    return {
      title: document.title,
      textContent: "",
      url: location.href,
    };
  }
}

/**
 * Fetch list item details for the current page via SharePoint REST API.
 * Returns augmented context if available, otherwise returns DOM extraction result.
 */
export async function getPageContext(
  spHttpClient: { get(url: string, factory: unknown, options?: unknown): Promise<{ ok: boolean; json(): Promise<unknown> }> },
  webUrl: string,
  spHttpClientConfiguration: unknown
): Promise<PageContext> {
  const domContext = extractPageContent();

  try {
    const relUrl = new URL(domContext.url).pathname;
    const apiUrl = `${webUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(relUrl)}')?$select=ListItemAllFields/Title,ListItemAllFields/Body&$expand=ListItemAllFields`;

    const response = await spHttpClient.get(apiUrl, spHttpClientConfiguration);
    if (!response.ok) {
      return domContext;
    }

    const data = await response.json() as {
      ListItemAllFields?: { Title?: string; Body?: string };
    };

    const restTitle = data.ListItemAllFields?.Title;
    const restBody = data.ListItemAllFields?.Body;

    if (restBody) {
      // Strip HTML from REST body
      const stripped = restBody.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
      return {
        ...domContext,
        title: restTitle || domContext.title,
        textContent: (stripped + " " + domContext.textContent).slice(0, 5000),
      };
    }
  } catch {
    // Silently fall back to DOM extraction
  }

  return domContext;
}
