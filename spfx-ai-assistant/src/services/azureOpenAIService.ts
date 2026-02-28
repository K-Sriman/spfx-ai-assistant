// ─── Azure OpenAI GPT-4o Service ─────────────────────────────────────────────
// All AI features are routed through this module.
// NOTE: API key embedded for enterprise SPFx prototype (corporate tenant-scoped).
// Production recommendation: proxy through Azure APIM with managed identity.

const CFG = {
  endpoint: "https://sharepoint-ai-openai.openai.azure.com",
  deployment: "gpt-4o",
  apiKey: process.env.AZURE_OPENAI_API_KEY,
  apiVersion: "2024-08-01-preview",
} as const;

export interface OAIMessage {
  role: "system" | "user" | "assistant";
  content: string;
}

export interface IntentResult {
  intent: "read_list" | "send_email" | "summarize_site" | "general_chat";
  listNameHint?: string;
  userQuestion?: string;

  recipientType?: "manager" | "specified" | "unspecified";

  emailTo?: string[];
  emailCc?: string[];
  emailContext?: string;
  dates?: string[];
  directResponse?: string;
}

export interface EmailDraftResult {
  subject: string;
  htmlBody: string;
  plainText: string;
}

// ─── Core fetch wrapper ───────────────────────────────────────────────────────

async function callGPT(
  messages: OAIMessage[],
  maxTokens = 2000,
  jsonMode = false
): Promise<string> {
  const url = `${CFG.endpoint}/openai/deployments/${CFG.deployment}/chat/completions?api-version=${CFG.apiVersion}`;

  const bodyPayload: Record<string, unknown> = {
    messages,
    max_tokens: maxTokens,
    temperature: jsonMode ? 0.1 : 0.7,
  };
  if (jsonMode) {
    bodyPayload.response_format = { type: "json_object" };
  }

  const res = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "api-key": CFG.apiKey,
    },
    body: JSON.stringify(bodyPayload),
  });

  if (!res.ok) {
    const errText = await res.text().catch(() => "");
    throw new Error(`OpenAI ${res.status}: ${errText.slice(0, 300)}`);
  }

  const data = await res.json() as { choices: Array<{ message: { content: string } }> };
  return (data.choices?.[0]?.message?.content ?? "").trim();
}

// ─── 1. Intent Detection ──────────────────────────────────────────────────────

export async function detectIntent(
  userMessage: string,
  history: OAIMessage[],
  siteName: string,
  availableLists: string[]
): Promise<IntentResult> {
  const listCtx = availableLists.length > 0
    ? `Available SharePoint lists on this site: ${availableLists.join(", ")}`
    : "SharePoint lists will be discovered on demand.";

const system = `You are an intent classifier for a SharePoint site called "${siteName}".
Analyze the user message and classify their intent.

${listCtx}

Return ONLY a valid JSON object:
{
  "intent": "read_list" | "send_email" | "summarize_site" | "general_chat",
  "listNameHint": "partial or full list name (read_list only)",
  "userQuestion": "specific question (read_list only)",
  "recipientType": "manager" | "specified" | "unspecified",
  "emailTo": ["email@domain.com"],        // only if explicit addresses or names are given
  "emailCc": [],
  "emailContext": "what the email should say (send_email only)",
  "dates": ["YYYY-MM-DD"],
  "directResponse": "text (general_chat only)"
}

Classification rules:
- "read_list": user asks about list contents, statuses, items, counts, etc.
- "send_email": user wants to draft/compose/send an email (leave requests, notices, etc.)
- "summarize_site": user asks to summarize the site/page
- "general_chat": anything else (greetings, help, etc.)

Recipient rules:
- If the user says "to my manager", "to manager", "to my boss/lead/line manager", set "recipientType": "manager".
- If the user provides explicit recipient(s) (emails or names), set "recipientType": "specified" and put them in "emailTo".
- If they don't specify who to send to, set "recipientType": "unspecified" and leave "emailTo": [].

Examples:

User: "Draft an email to my manager about WFH tomorrow."
{
  "intent": "send_email",
  "recipientType": "manager",
  "emailTo": [],
  "emailCc": [],
  "emailContext": "Draft a message requesting WFH tomorrow to the manager",
  "dates": []
}

User: "Email ramesh.k@contoso.com the weekly status"
{
  "intent": "send_email",
  "recipientType": "specified",
  "emailTo": ["ramesh.k@contoso.com"],
  "emailCc": [],
  "emailContext": "Weekly status update",
  "dates": []
}

User: "Write a project update email"
{
  "intent": "send_email",
  "recipientType": "unspecified",
  "emailTo": [],
  "emailCc": [],
  "emailContext": "Project update",
  "dates": []
}`;

  const messages: OAIMessage[] = [
    { role: "system", content: system },
    ...history.slice(-6),
    { role: "user", content: userMessage },
  ];

  const raw = await callGPT(messages, 600, true);
  try {
    return JSON.parse(raw) as IntentResult;
  } catch {
    return { intent: "general_chat", directResponse: raw };
  }
}

// ─── 2. List Q&A ──────────────────────────────────────────────────────────────

export async function answerFromListData(params: {
  userQuestion: string;
  listTitle: string;
  columns: string[];
  items: Record<string, unknown>[];
  history: OAIMessage[];
}): Promise<string> {
  const { userQuestion, listTitle, columns, items, history } = params;
  const sample = items.slice(0, 60);
  const dataJson = JSON.stringify({ columns, items: sample }).slice(0, 9000);

  const system = `You are a SharePoint data analyst. The user asked about the "${listTitle}" list.

List data (showing ${sample.length} of ${items.length} items):
${dataJson}

${items.length > 60 ? `⚠️ Showing first 60 of ${items.length} total items.` : ""}

Answer the user's question accurately based on the data above.
Format your response professionally:
- Start with a direct answer
- Use **bold** for key values
- Use bullet lists (- item) for multiple items
- Use ## for section headers when needed
- Provide counts and summaries where relevant
- If grouping by status or category, show breakdown
- Never invent data that is not in the list`;

  const msgs: OAIMessage[] = [
    { role: "system", content: system },
    ...history.slice(-4),
    { role: "user", content: userQuestion },
  ];

  return callGPT(msgs, 1500);
}

// ─── 3. Email Drafting ────────────────────────────────────────────────────────

export async function draftEmail(params: {
  context: string;
  fromName: string;
  fromTitle?: string;
  fromEmail?: string;
  fromDepartment?: string;
  toEmails: string[];
  ccEmails: string[];
  managerName?: string;
  managerEmail?: string;
  dates?: string[];
  history: OAIMessage[];
}): Promise<EmailDraftResult> {
  const { context, fromName, toEmails, ccEmails, managerName, dates } = params;
  const isLeave = /leave|vacation|day off|holiday|absent|pto|ooo|time off/i.test(context);

  const recipientInfo = toEmails.length > 0
    ? `To: ${toEmails.join(", ")}`
    : managerName
    ? `To: ${managerName} (sender's manager)`
    : "To: (will be specified)";

  const datesInfo = dates && dates.length > 0
    ? `Dates: ${dates.join(", ")}`
    : "";

const system = `You are a professional email writer for a corporate environment.
Draft a ${isLeave ? "leave request" : "professional"} email.

Sender Info:
- Name: ${fromName}
- Title: ${params.fromTitle ?? ""}
- Email: ${params.fromEmail ?? ""}
- Department: ${params.fromDepartment ?? ""}

Recipients:
${recipientInfo}
${ccEmails.length ? `CC: ${ccEmails.join(", ")}` : ""}

${datesInfo}

Context/Request: ${context}

SIGNATURE REQUIREMENTS:
- DO NOT use placeholders like [Your Name].
- Use the sender info from above.
- Must include:
  - Name
  - Job Title (if available)
  - Department (if available)
  - Email address

Return ONLY valid JSON:
{
  "subject": "...",
  "htmlBody": "...",
  "plainText": "..."
}
From: ${fromName}
${recipientInfo}
${ccEmails.length ? `CC: ${ccEmails.join(", ")}` : ""}
${datesInfo}
Context/Request: ${context}

Return ONLY valid JSON:
{
  "subject": "concise professional subject line",
  "htmlBody": "full HTML email with inline styles (Segoe UI 14px), proper greeting, well-structured body paragraphs, and professional closing with sender's name",
  "plainText": "plain text version — same content as htmlBody, no HTML tags"
}

Requirements:
- Professional, polished tone appropriate for corporate email
- Clear structure: greeting → purpose → details → request/action → closing → signature
- For leave: mention specific dates, brief reason, assurance of handover/coverage
- Concise — not overly long
- htmlBody must be complete valid HTML with inline styles for email client compatibility`;

  const msgs: OAIMessage[] = [
    { role: "system", content: system },
    ...params.history.slice(-4),
    { role: "user", content: `Write an email about: ${context}` },
  ];

  const raw = await callGPT(msgs, 1500, true);
  try {
    const parsed = JSON.parse(raw) as EmailDraftResult;
    return parsed;
  } catch {
    return {
      subject: isLeave ? "Leave Request" : "Email",
      htmlBody: `<p style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;">${raw}</p>`,
      plainText: raw,
    };
  }
}

// ─── 4. Site Summarization ────────────────────────────────────────────────────

export async function summarizeSite(params: {
  siteName: string;
  siteUrl: string;
  pageTitle: string;
  pageContent: string;
  listNames?: string[];
  history: OAIMessage[];
}): Promise<string> {
  const { siteName, siteUrl, pageTitle, pageContent, listNames, history } = params;

  const system = `You are a SharePoint site assistant. Provide a clear, insightful summary.

Site: ${siteName}
URL: ${siteUrl}
Current Page: ${pageTitle}
${listNames && listNames.length > 0 ? `Lists on this site: ${listNames.join(", ")}` : ""}

Page content extracted from DOM:
"""
${pageContent.slice(0, 5000)}
"""

Provide a well-structured, professional summary using this format:

## Overview
[2–3 sentences describing what this site/page is for and who it serves]

## What's on This Page
[Key content and information found on the current page]

## Site Structure
[Based on available lists and navigation — what capabilities/sections this site has]

## Key Highlights
[3–5 bullet points of the most important or recent content]

Be factual — only describe what is clearly present in the content. Do not invent information.
Use **bold** for important terms. Keep it professional and scannable.`;

  const msgs: OAIMessage[] = [
    { role: "system", content: system },
    ...history.slice(-2),
    { role: "user", content: "Please summarize this SharePoint site for me." },
  ];

  return callGPT(msgs, 1200);
}

// ─── 5. General Chat ──────────────────────────────────────────────────────────

export async function generalChat(params: {
  userMessage: string;
  siteName: string;
  history: OAIMessage[];
  directResponse?: string;
}): Promise<string> {
  if (params.directResponse && params.directResponse.length > 20) {
    return params.directResponse;
  }

  const system = `You are a helpful AI assistant embedded in a SharePoint site called "${params.siteName}".
Help users with their questions. Be concise, friendly, and professional.
Capabilities you can tell users about:
- Find and query data from any SharePoint list on this site
- Draft and send professional emails (including leave requests)
- Summarize the site or current page content
- Answer general workplace or Microsoft 365 questions`;

  const msgs: OAIMessage[] = [
    { role: "system", content: system },
    ...params.history.slice(-8),
    { role: "user", content: params.userMessage },
  ];

  return callGPT(msgs, 800);
}
