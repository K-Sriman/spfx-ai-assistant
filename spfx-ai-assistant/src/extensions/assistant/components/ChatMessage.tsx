import * as React from "react";
import type { ChatEntry, EmailDraftData } from "../../../services/types";
import styles from "./styles.module.scss";
import { textToHtml } from "../../../utils/formatting";

// ─── Markdown Renderer ────────────────────────────────────────────────────────

function parseInline(text: string): React.ReactNode[] {
  // Handle **bold**, *italic*, `code`
  const parts = text.split(/(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)/g);
  return parts.map((part, i) => {
    if (part.startsWith("**") && part.endsWith("**"))
      return <strong key={i}>{part.slice(2, -2)}</strong>;
    if (part.startsWith("*") && part.endsWith("*"))
      return <em key={i}>{part.slice(1, -1)}</em>;
    if (part.startsWith("`") && part.endsWith("`"))
      return <code key={i}>{part.slice(1, -1)}</code>;
    return part;
  });
}

function MarkdownContent({ text }: { text: string }): JSX.Element {
  const lines = text.split("\n");
  const elements: React.ReactNode[] = [];
  const pendingList: { key: number; tag: "ul" | "ol"; items: React.ReactNode[] }[] = [];

  function flushList(): void {
    if (pendingList.length > 0) {
      const { key, tag: Tag, items } = pendingList[pendingList.length - 1];
      elements.push(
        <Tag key={`list-${key}`}>
          {items}
        </Tag>
      );
      pendingList.pop();
    }
  }

  function currentList(): { key: number; tag: "ul" | "ol"; items: React.ReactNode[] } | null {
    return pendingList.length > 0 ? pendingList[pendingList.length - 1] : null;
  }

  lines.forEach((raw, i) => {
    const line = raw;

    // Horizontal rule
    if (/^---+$/.test(line.trim())) {
      flushList();
      elements.push(<hr key={i} />);
      return;
    }

    // Headers
    if (line.startsWith("#### ")) {
      flushList();
      elements.push(<h4 key={i}>{parseInline(line.slice(5))}</h4>);
      return;
    }
    if (line.startsWith("### ")) {
      flushList();
      elements.push(<h3 key={i}>{parseInline(line.slice(4))}</h3>);
      return;
    }
    if (line.startsWith("## ")) {
      flushList();
      elements.push(<h2 key={i}>{parseInline(line.slice(3))}</h2>);
      return;
    }
    if (line.startsWith("# ")) {
      flushList();
      elements.push(<h2 key={i}>{parseInline(line.slice(2))}</h2>);
      return;
    }

    // Unordered list
    if (/^[-*•] /.test(line)) {
      const cl = currentList();
      if (cl && cl.tag === "ul") {
        cl.items.push(<li key={i}>{parseInline(line.slice(2))}</li>);
      } else {
        flushList();
        pendingList.push({ key: i, tag: "ul", items: [<li key={i}>{parseInline(line.slice(2))}</li>] });
      }
      return;
    }

    // Ordered list
    if (/^\d+\. /.test(line)) {
      const text2 = line.replace(/^\d+\. /, "");
      const cl = currentList();
      if (cl && cl.tag === "ol") {
        cl.items.push(<li key={i}>{parseInline(text2)}</li>);
      } else {
        flushList();
        pendingList.push({ key: i, tag: "ol", items: [<li key={i}>{parseInline(text2)}</li>] });
      }
      return;
    }

    // Blank line
    if (line.trim() === "") {
      flushList();
      return;
    }

    // Normal paragraph
    flushList();
    elements.push(<p key={i}>{parseInline(line)}</p>);
  });

  flushList();

  return <div className={styles.mdContent}>{elements}</div>;
}

// ─── Typing indicator ─────────────────────────────────────────────────────────

export function TypingIndicator(): JSX.Element {
  return (
    <div className={styles.typingRow}>
      <div className={`${styles.avatar} ${styles.aiAvatar}`}>✦</div>
      <div className={styles.typingBubble} aria-label="AI is thinking">
        <span className={styles.typingDot} />
        <span className={styles.typingDot} />
        <span className={styles.typingDot} />
      </div>
    </div>
  );
}

// ─── Email draft card ─────────────────────────────────────────────────────────

interface EmailCardProps {
  data: EmailDraftData;
  onConfirm: () => void;
  onCancel: () => void;
  onUpdate: (updated: Partial<EmailDraftData>) => void;
}

function EmailCard({ data, onConfirm, onCancel, onUpdate }: EmailCardProps): JSX.Element {
  const [editing, setEditing] = React.useState(false);
  const [editTo, setEditTo] = React.useState(data.to.join(", "));
  const [editCc, setEditCc] = React.useState(data.cc.join(", "));
  const [editSubject, setEditSubject] = React.useState(data.subject);
  const [editBody, setEditBody] = React.useState(data.plainText);

  if (data.sent) {
    return (
      <div className={styles.successCard}>
        <div className={styles.successIcon}>✅</div>
        <div className={styles.successTitle}>Email Sent Successfully!</div>
        <div className={styles.successSub}>Your email has been delivered.</div>
      </div>
    );
  }

  if (data.cancelled) {
    return (
      <div className={styles.aiBubble} style={{ padding: "10px 14px", fontStyle: "italic", color: "#6b7280", fontSize: 12 }}>
        Email draft cancelled.
      </div>
    );
  }


function saveEdit(): void {
  onUpdate({
    to: editTo.split(",").map(s => s.trim()).filter(Boolean),
    cc: editCc.split(",").map(s => s.trim()).filter(Boolean),
    subject: editSubject,
    plainText: editBody,
    htmlBody: textToHtml(editBody), // <-- use robust conversion here
  });
  setEditing(false);
}


  return (
    <div className={styles.emailCard}>
      <div className={styles.emailCardHeader}>
        <span>📧</span>
        <span>Draft Email</span>
        {data.loading && <span style={{ marginLeft: "auto", opacity: 0.8 }}>Sending…</span>}
      </div>

      <div className={styles.emailCardBody}>
        {editing ? (
          <>
            <div className={styles.emailField}>
              <label className={styles.emailLabel}>To</label>
              <input
                className={styles.emailEditInput}
                value={editTo}
                onChange={e => setEditTo(e.target.value)}
                placeholder="recipient@company.com"
              />
            </div>
            <div className={styles.emailField}>
              <label className={styles.emailLabel}>CC</label>
              <input
                className={styles.emailEditInput}
                value={editCc}
                onChange={e => setEditCc(e.target.value)}
                placeholder="cc@company.com (optional)"
              />
            </div>
            <div className={styles.emailField}>
              <label className={styles.emailLabel}>Subject</label>
              <input
                className={styles.emailEditInput}
                value={editSubject}
                onChange={e => setEditSubject(e.target.value)}
              />
            </div>
            <div className={styles.emailField}>
              <label className={styles.emailLabel}>Body</label>
              <textarea
                className={styles.emailEditTextarea}
                value={editBody}
                onChange={e => setEditBody(e.target.value)}
              />
            </div>
          </>
        ) : (
          <>
            <div className={styles.emailField}>
              <div className={styles.emailLabel}>To</div>
              <div className={styles.emailValue}>
                {data.to.length > 0 ? data.to.join(", ") : (
                  <span style={{ color: "#ef4444", fontStyle: "italic" }}>No recipient — please edit</span>
                )}
              </div>
            </div>
            {data.cc.length > 0 && (
              <div className={styles.emailField}>
                <div className={styles.emailLabel}>CC</div>
                <div className={styles.emailValue}>{data.cc.join(", ")}</div>
              </div>
            )}
            <div className={styles.emailField}>
              <div className={styles.emailLabel}>Subject</div>
              <div className={styles.emailValue}>{data.subject}</div>
            </div>
            <div className={styles.emailField}>
              <div className={styles.emailLabel}>Body</div>
              <div className={styles.emailBody}>{data.plainText}</div>
            </div>
          </>
        )}
      </div>

      <div className={styles.emailActions}>
        {editing ? (
          <>
            <button className={styles.btnPrimary} onClick={saveEdit}>Save Changes</button>
            <button className={styles.btnOutline} onClick={() => setEditing(false)}>Cancel Edit</button>
          </>
        ) : (
          <>
            <button
              className={styles.btnPrimary}
              onClick={onConfirm}
              disabled={data.loading || data.to.length === 0}
            >
              {data.loading ? "Sending…" : "✔ Confirm & Send"}
            </button>
            <button
              className={styles.btnOutline}
              onClick={() => setEditing(true)}
              disabled={data.loading}
            >
              ✏ Edit
            </button>
            <button
              className={styles.btnDanger}
              onClick={onCancel}
              disabled={data.loading}
            >
              ✕ Cancel
            </button>
          </>
        )}
      </div>
    </div>
  );
}

// ─── Main ChatMessage component ───────────────────────────────────────────────

interface ChatMessageProps {
  entry: ChatEntry;
  userInitials: string;
  onEmailConfirm: (id: string) => void;
  onEmailCancel: (id: string) => void;
  onEmailUpdate: (id: string, updated: Partial<EmailDraftData>) => void;
}

function formatTime(ts: number): string {
  return new Date(ts).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
}

const ChatMessage: React.FC<ChatMessageProps> = ({
  entry,
  userInitials,
  onEmailConfirm,
  onEmailCancel,
  onEmailUpdate,
}) => {
  // Thinking state
  if (entry.type === "thinking") {
    return <TypingIndicator />;
  }

  // Error
  if (entry.type === "error") {
    return (
      <div className={styles.errorBubble} role="alert">
        <span>⚠️</span>
        <span>{entry.text}</span>
      </div>
    );
  }

  // User message
  if (entry.role === "user") {
    return (
      <div className={`${styles.msgRow} ${styles.userRow}`}>
        <div className={`${styles.avatar} ${styles.userAvatar}`}>{userInitials}</div>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end" }}>
          <div className={`${styles.bubble} ${styles.userBubble}`}>{entry.text}</div>
          <span className={`${styles.msgTime} ${styles.userTime}`}>{formatTime(entry.timestamp)}</span>
        </div>
      </div>
    );
  }

  // Email draft
  if (entry.type === "email_draft" && entry.emailData) {
    return (
      <div className={styles.msgRow}>
        <div className={`${styles.avatar} ${styles.aiAvatar}`}>✦</div>
        <div style={{ flex: 1, minWidth: 0 }}>
          <EmailCard
            data={entry.emailData}
            onConfirm={() => onEmailConfirm(entry.id)}
            onCancel={() => onEmailCancel(entry.id)}
            onUpdate={(u) => onEmailUpdate(entry.id, u)}
          />
          <span className={styles.msgTime}>{formatTime(entry.timestamp)}</span>
        </div>
      </div>
    );
  }

  // AI text message
  return (
    <div className={styles.msgRow}>
      <div className={`${styles.avatar} ${styles.aiAvatar}`}>✦</div>
      <div style={{ display: "flex", flexDirection: "column", minWidth: 0, flex: 1 }}>
        <div className={`${styles.bubble} ${styles.aiBubble}`}>
          <MarkdownContent text={entry.text} />
        </div>
        <span className={styles.msgTime}>{formatTime(entry.timestamp)}</span>
      </div>
    </div>
  );
};

export default ChatMessage;
