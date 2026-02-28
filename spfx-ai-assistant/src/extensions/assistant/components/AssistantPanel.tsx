import * as React from "react";
import type { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { SPHttpClient } from "@microsoft/sp-http";
import type { ChatEntry, EmailDraftData, ManagerProfile, UserProfile } from "../../../services/types";
import type { OAIMessage } from "../../../services/azureOpenAIService";
import {
  detectIntent,
  answerFromListData,
  draftEmail,
  summarizeSite,
  generalChat,
} from "../../../services/azureOpenAIService";
import {
  getAllLists,
  findListByHint,
  getListItems,
  type ListInfo,
} from "../../../services/spListService";
import { getMe, getManager } from "../../../services/graphUserService";
import { sendMail } from "../../../services/graphMailService";
import { extractPageContent } from "../../../services/spPageService";
import { extractAndResolveDates, formatDateHuman } from "../../../utils/formatting";
import { trackEvent, trackError } from "../../../utils/telemetry";
import ChatMessage, { TypingIndicator } from "./ChatMessage";
import styles from "./styles.module.scss";

// ─── Session Storage helpers ──────────────────────────────────────────────────
const SESSION_KEY = "ai-assistant-chat-v2";

function loadHistory(): ChatEntry[] {
  try {
    const raw = sessionStorage.getItem(SESSION_KEY);
    return raw ? (JSON.parse(raw) as ChatEntry[]) : [];
  } catch {
    return [];
  }
}

function saveHistory(entries: ChatEntry[]): void {
  try {
    sessionStorage.setItem(SESSION_KEY, JSON.stringify(entries.slice(-30)));
  } catch { /* quota or unavailable */ }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function toOAIHistory(entries: ChatEntry[]): OAIMessage[] {
  return entries
    .filter((e) => e.type === "text")
    .map((e) => ({ role: e.role as "user" | "assistant", content: e.text }));
}

function uid(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

// ─── Static suggestions ───────────────────────────────────────────────────────
const SUGGESTIONS = [
  { icon: "📋", text: "What's are the list present ?" },
  { icon: "📄", text: "Summarize this site" },
  { icon: "📧", text: "Apply leave for tomorrow and day after tomorrow" },
  { icon: "🔍", text: "Show me recent documents" },
  { icon: "✉️", text: "Draft an email to my manager about the project update" },
  { icon: "📊", text: "List all items with their status" },
];

// ─── Props ────────────────────────────────────────────────────────────────────
interface AssistantPanelProps {
  spContext: ApplicationCustomizerContext;
  onClose: () => void;
}

// ─── Component ────────────────────────────────────────────────────────────────
const AssistantPanel: React.FC<AssistantPanelProps> = ({ spContext, onClose }) => {
  const [messages, setMessages] = React.useState<ChatEntry[]>(loadHistory);
  const [inputText, setInputText] = React.useState("");
  const [isLoading, setIsLoading] = React.useState(false);
  const [availableLists, setAvailableLists] = React.useState<ListInfo[]>([]);
  const [userInitials, setUserInitials] = React.useState("U");
  const [mode, setMode] = React.useState<"site" | "web">("site");

  const chatEndRef = React.useRef<HTMLDivElement>(null);
  const inputRef = React.useRef<HTMLTextAreaElement>(null);
  const listsLoadedRef = React.useRef(false);

  const siteName = spContext.pageContext.web.title;
  const siteUrl = spContext.pageContext.web.absoluteUrl;

  // ── Init: fetch user profile + lists ───────────────────────────────────────
  React.useEffect(() => {
    getMe()
      .then((me) => {
        const parts = me.displayName.split(" ");
        const initials =
          parts.length >= 2
            ? (parts[0][0] + parts[parts.length - 1][0]).toUpperCase()
            : me.displayName.slice(0, 2).toUpperCase();
        setUserInitials(initials);
      })
      .catch(() => void 0);

    loadLists().catch(() => void 0);

    return void 0;
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Auto-scroll on new message
  React.useEffect(() => {
    chatEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  // Persist to sessionStorage
  React.useEffect(() => {
    saveHistory(messages.filter((m) => m.type !== "thinking"));
  }, [messages]);

  // ── Load all lists ─────────────────────────────────────────────────────────
  async function loadLists(): Promise<ListInfo[]> {
    if (listsLoadedRef.current) return availableLists;
    const spClient = spContext.spHttpClient as SPHttpClient;
    const lists = await getAllLists(spClient, siteUrl);
    setAvailableLists(lists);
    listsLoadedRef.current = true;
    return lists;
  }

  // ── Message helpers ────────────────────────────────────────────────────────
  function addEntry(entry: Omit<ChatEntry, "id" | "timestamp">): string {
    const id = uid();
    setMessages((prev) => [...prev, { ...entry, id, timestamp: Date.now() }]);
    return id;
  }

  function replaceEntry(id: string, update: Partial<ChatEntry>): void {
    setMessages((prev) => prev.map((m) => (m.id === id ? { ...m, ...update } : m)));
  }

  function removeEntry(id: string): void {
    setMessages((prev) => prev.filter((m) => m.id !== id));
  }

  function updateEmailData(id: string, update: Partial<EmailDraftData>): void {
    setMessages((prev) =>
      prev.map((m) =>
        m.id === id && m.emailData
          ? { ...m, emailData: { ...m.emailData, ...update } }
          : m
      )
    );
  }

  // ── Main send handler ──────────────────────────────────────────────────────
  async function handleSend(text: string): Promise<void> {
    const query = text.trim();
    if (!query || isLoading) return;

    setInputText("");
    setIsLoading(true);
    trackEvent({ name: "query_submitted", properties: { mode } });

    addEntry({ role: "user", text: query, type: "text" });

    // Thinking bubble
    const thinkId = uid();
    setMessages((prev) => [
      ...prev,
      { id: thinkId, role: "assistant", text: "", type: "thinking", timestamp: Date.now() },
    ]);

    try {
      const lists = await loadLists();
      const history = toOAIHistory(messages);

      // Detect intent via GPT-4o
      const intent = await detectIntent(
        query,
        history,
        siteName,
        lists.map((l) => l.Title)
      );

      removeEntry(thinkId);

      switch (intent.intent) {
        case "read_list":
          await handleReadList(
            intent.listNameHint ?? "",
            intent.userQuestion ?? query,
            lists,
            history
          );
          break;

       case "send_email":
          await handleDraftEmail(
            intent.emailContext ?? query,
            intent.emailTo ?? [],
            intent.emailCc ?? [],
            intent.dates ?? extractAndResolveDates(query),
            history,
            intent.recipientType // NEW: pass the recipient type
          );
          break;

        case "summarize_site":
          await handleSummarizeSite(history);
          break;

        default: {
          const response = await generalChat({
            userMessage: query,
            siteName,
            history,
            directResponse: intent.directResponse,
          });
          addEntry({ role: "assistant", text: response, type: "text" });
        }
      }
    } catch (err) {
      removeEntry(thinkId);
      const msg =
        err instanceof Error ? err.message : "Something went wrong. Please try again.";
      addEntry({ role: "assistant", text: msg, type: "error" });
      trackError("AssistantPanel.handleSend", err);
    } finally {
      setIsLoading(false);
    }
  }

  // ── Read SharePoint list ────────────────────────────────────────────────────
  async function handleReadList(
    listHint: string,
    userQuestion: string,
    lists: ListInfo[],
    history: OAIMessage[]
  ): Promise<void> {
    const matched = findListByHint(lists, listHint);

    if (!matched) {
      const names = lists.map((l) => `**${l.Title}**`).join(", ") || "_None discovered_";
      addEntry({
        role: "assistant",
        text: `I couldn't find a list matching "${listHint}" on this site.\n\nAvailable lists: ${names}\n\nTry asking about one of those by name.`,
        type: "text",
      });
      return;
    }

    // Show fetching status
    const fetchId = uid();
    setMessages((prev) => [
      ...prev,
      {
        id: fetchId,
        role: "assistant",
        text: `Reading **${matched.Title}** (${matched.ItemCount} items)…`,
        type: "text",
        timestamp: Date.now(),
      },
    ]);

    try {
      const spClient = spContext.spHttpClient as SPHttpClient;
      const result = await getListItems(spClient, siteUrl, matched.Title, []);

      const answer = await answerFromListData({
        userQuestion,
        listTitle: result.listTitle,
        columns: result.columns,
        items: result.items,
        history,
      });

      replaceEntry(fetchId, { text: answer, type: "text" });
      trackEvent({ name: "knowledge_search_complete", properties: { list: matched.Title } });
    } catch (err) {
      const msg = err instanceof Error ? err.message : "Failed to read list data.";
      replaceEntry(fetchId, { text: `⚠️ ${msg}`, type: "error" });
      trackError("AssistantPanel.handleReadList", err);
    }
  }

  // ── Draft email ─────────────────────────────────────────────────────────────

async function handleDraftEmail(
  context: string,
  toEmails: string[],
  ccEmails: string[],
  dates: string[],
  history: OAIMessage[],
  recipientType?: "manager" | "specified" | "unspecified" // NEW param
): Promise<void> {
  try {
    const [me, manager]: [UserProfile, ManagerProfile | null] = await Promise.all([
  getMe().catch(() => ({
    displayName: "You",
    mail: "",
    jobTitle: undefined,
    department: undefined,
  } as UserProfile)),
  getManager().catch(() => null),
]);

    // Only default to manager if user explicitly asked to send to manager
    const resolvedTo =
      toEmails.length > 0
        ? toEmails
        : (recipientType === "manager" && manager?.mail)
        ? [manager.mail]
        : [];

    const humanDates = dates.map(formatDateHuman);
    const datesCtx = humanDates.length > 0 ? ` for ${humanDates.join(" and ")}` : "";
    const fullContext = context + datesCtx;

  const drafted = await draftEmail({
  context: fullContext,
  fromName: me.displayName,
  fromTitle: me.jobTitle,
  fromEmail: me.mail,
  fromDepartment: me.department,
  toEmails: resolvedTo,
  ccEmails,
  managerName: manager?.displayName,
  managerEmail: manager?.mail,
  dates: humanDates,
  history,
});

    const emailData: EmailDraftData = {
      to: resolvedTo,
      cc: ccEmails,
      subject: drafted.subject,
      plainText: drafted.plainText,
      htmlBody: drafted.htmlBody,
      sent: false,
      cancelled: false,
      loading: false,
    };

    addEntry({
      role: "assistant",
      text: `I've drafted a professional email for you. Please review the details below — you can edit any field before sending.`,
      type: "text",
    });
    addEntry({ role: "assistant", text: "", type: "email_draft", emailData });
    trackEvent({ name: "email_draft_shown" });
  } catch (err) {
    const msg = err instanceof Error ? err.message : "Could not draft email.";
    addEntry({ role: "assistant", text: `⚠️ ${msg}`, type: "error" });
    trackError("AssistantPanel.handleDraftEmail", err);
  }
}

  // ── Summarize site ──────────────────────────────────────────────────────────
  async function handleSummarizeSite(history: OAIMessage[]): Promise<void> {
    try {
      const pageCtx = extractPageContent();
      const lists = await loadLists();

      const summary = await summarizeSite({
        siteName,
        siteUrl,
        pageTitle: pageCtx.title,
        pageContent: pageCtx.textContent,
        listNames: lists.map((l) => l.Title),
        history,
      });

      addEntry({ role: "assistant", text: summary, type: "text" });
      trackEvent({ name: "knowledge_search_complete", properties: { action: "summarize" } });
    } catch (err) {
      const msg = err instanceof Error ? err.message : "Could not summarize site.";
      addEntry({ role: "assistant", text: `⚠️ ${msg}`, type: "error" });
      trackError("AssistantPanel.handleSummarizeSite", err);
    }
  }

  // ── Email confirm / cancel ──────────────────────────────────────────────────
  async function handleEmailConfirm(id: string): Promise<void> {
    const entry = messages.find((m) => m.id === id);
    if (!entry?.emailData) return;

    updateEmailData(id, { loading: true });
    try {
      await sendMail({
        to: entry.emailData.to,
        cc: entry.emailData.cc,
        subject: entry.emailData.subject,
        htmlBody: entry.emailData.htmlBody,
      });
      updateEmailData(id, { sent: true, loading: false });
      addEntry({ role: "assistant", text: "✅ Your email has been sent successfully!", type: "text" });
      trackEvent({ name: "email_sent" });
    } catch (err) {
      updateEmailData(id, { loading: false });
      const msg = err instanceof Error ? err.message : "Send failed.";
      addEntry({
        role: "assistant",
        text: `⚠️ Could not send email: ${msg}\n\nEnsure **Mail.Send** permission is approved in SharePoint Admin → API Access.`,
        type: "error",
      });
      trackError("AssistantPanel.handleEmailConfirm", err);
    }
  }

  function handleEmailCancel(id: string): void {
    updateEmailData(id, { cancelled: true });
    addEntry({
      role: "assistant",
      text: "Email draft cancelled. Let me know if you need anything else.",
      type: "text",
    });
    trackEvent({ name: "email_cancelled" });
  }

  // ── Input auto-grow ─────────────────────────────────────────────────────────
  function handleInputChange(e: React.ChangeEvent<HTMLTextAreaElement>): void {
    setInputText(e.target.value);
    const el = e.target;
    el.style.height = "auto";
    el.style.height = `${Math.min(el.scrollHeight, 110)}px`;
  }

  function handleKeyDown(e: React.KeyboardEvent<HTMLTextAreaElement>): void {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      void handleSend(inputText);
    }
  }

  // ── Render ──────────────────────────────────────────────────────────────────
  return (
    <div className={styles.panel} role="dialog" aria-label="AI Assistant" aria-modal="true">

      {/* ── Header ── */}
      <div className={styles.header}>
        <div className={styles.headerTop}>
          <div className={styles.aiLogo} aria-hidden="true">✦</div>
          <div className={styles.headerInfo}>
            <div className={styles.headerTitle}>{siteName}</div>
            <div className={styles.headerSubtitle}>AI Assistant · GPT-4o</div>
          </div>
          <button className={styles.closeBtn} onClick={onClose} aria-label="Close assistant">
            ✕
          </button>
        </div>
        <div className={styles.modeRow}>
          <button
            className={`${styles.modeChip} ${mode === "site" ? styles.modeChipActive : ""}`}
            onClick={() => setMode("site")}
            aria-pressed={mode === "site"}
          >
            🏢 This Site
          </button>
          <button
            className={`${styles.modeChip} ${mode === "web" ? styles.modeChipActive : ""}`}
            onClick={() => setMode("web")}
            aria-pressed={mode === "web"}
          >
            🌐 General
          </button>
        </div>
      </div>

      {/* ── Chat area ── */}
      <div className={styles.chatArea} role="log" aria-live="polite" aria-atomic="false">

        {/* Welcome / suggestion screen */}
        {messages.length === 0 && (
          <div className={styles.welcome}>
            <div className={styles.welcomeLogo} aria-hidden="true">✦</div>
            <div className={styles.welcomeTitle}>How can I help you today?</div>
            <div className={styles.welcomeText}>
              I can search any SharePoint list, draft professional emails, summarize
              this site, and answer questions — all powered by GPT-4o.
            </div>
            <div className={styles.suggestionGrid}>
              {SUGGESTIONS.map((s, i) => (
                <button
                  key={i}
                  className={styles.suggestionCard}
                  onClick={() => void handleSend(s.text)}
                  aria-label={`Try: ${s.text}`}
                >
                  <span className={styles.suggestionIcon}>{s.icon}</span>
                  <span className={styles.suggestionText}>{s.text}</span>
                </button>
              ))}
            </div>
          </div>
        )}

        {/* Message list */}
        {messages.map((entry) =>
          entry.type === "thinking" ? (
            <TypingIndicator key={entry.id} />
          ) : (
            <ChatMessage
              key={entry.id}
              entry={entry}
              userInitials={userInitials}
              onEmailConfirm={(id) => void handleEmailConfirm(id)}
              onEmailCancel={handleEmailCancel}
              onEmailUpdate={updateEmailData}
            />
          )
        )}

        <div ref={chatEndRef} aria-hidden="true" />
      </div>

      {/* ── Input area ── */}
      <div className={styles.inputArea}>
        <div className={styles.inputHint}>
          Press Enter to send · Shift+Enter for new line
        </div>
        <div className={styles.inputRow}>
          <textarea
            ref={inputRef}
            className={styles.chatInput}
            value={inputText}
            onChange={handleInputChange}
            onKeyDown={handleKeyDown}
            placeholder="Ask about lists, request leave, summarize site…"
            rows={1}
            disabled={isLoading}
            aria-label="Message input"
          />
          <button
            className={styles.sendBtn}
            onClick={() => void handleSend(inputText)}
            disabled={!inputText.trim() || isLoading}
            aria-label="Send message"
          >
            {/* Send arrow SVG */}
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"
              strokeLinecap="round" strokeLinejoin="round">
              <line x1="22" y1="2" x2="11" y2="13" />
              <polygon points="22 2 15 22 11 13 2 9 22 2" />
            </svg>
          </button>
        </div>
      </div>
    </div>
  );
};

export default AssistantPanel;
