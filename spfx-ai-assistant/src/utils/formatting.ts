// ─── Date / Text Formatting Utilities ──────────────────────────────────────

/**
 * Resolve natural-language date phrases to ISO date strings (YYYY-MM-DD).
 */
export function resolveNaturalDate(phrase: string): string | null {
  const now = new Date();
  const lower = phrase.toLowerCase().trim();

  const offsetMap: Record<string, number> = {
    today: 0,
    tomorrow: 1,
    "day after tomorrow": 2,
    "day after tomorow": 2, // typo tolerance
    yesterday: -1,
    "next monday": daysUntilWeekday(now, 1),
    "next tuesday": daysUntilWeekday(now, 2),
    "next wednesday": daysUntilWeekday(now, 3),
    "next thursday": daysUntilWeekday(now, 4),
    "next friday": daysUntilWeekday(now, 5),
  };

  if (lower in offsetMap) {
    const d = new Date(now);
    d.setDate(d.getDate() + offsetMap[lower]);
    return toISO(d);
  }

  // Try parsing explicitly formatted dates
  const parsed = Date.parse(phrase);
  if (!isNaN(parsed)) {
    return toISO(new Date(parsed));
  }

  return null;
}

function daysUntilWeekday(from: Date, weekday: number): number {
  const current = from.getDay();
  const diff = (weekday - current + 7) % 7;
  return diff === 0 ? 7 : diff;
}

export function toISO(date: Date): string {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

/**
 * Format ISO date string to human-readable: "18 Feb 2025"
 */
export function formatDateHuman(iso: string): string {
  const months = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
  ];
  const [year, month, day] = iso.split("-").map(Number);
  return `${day} ${months[month - 1]} ${year}`;
}

/**
 * Extract all date-like phrases from a string and resolve them.
 */
export function extractAndResolveDates(input: string): string[] {
  const patterns = [
    /\btoday\b/gi,
    /\btomorrow\b/gi,
    /\bday after tomorrow\b/gi,
    /\byesterday\b/gi,
    /\bnext (?:monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/gi,
    /\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b/g,
    /\b\d{1,2}\s+(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\b/gi,
  ];

  const found = new Set<string>();
  for (const pattern of patterns) {
    const matches = input.match(pattern) || [];
    for (const match of matches) {
      const resolved = resolveNaturalDate(match);
      if (resolved) found.add(resolved);
    }
  }

  return Array.from(found).sort();
}

/**
 * Truncate text to a max number of words.
 */
export function truncateWords(text: string, maxWords: number): string {
  const words = text.trim().split(/\s+/);
  if (words.length <= maxWords) return text;
  return words.slice(0, maxWords).join(" ") + "…";
}

// ─── Email Text → HTML Formatting Utilities ──────────────────────────────────

/**
 * Escape HTML special characters in user-provided plain text.
 * Prevents accidental HTML injection inside the edited body.
 */
export function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/**
 * Converts user-edited plain text into robust email HTML (<p> and <br>)
 * that renders correctly across Outlook/OWA/Gmail.
 *
 * Rules:
 * - Paragraphs are separated by one or more blank lines
 * - Single newlines within a paragraph become <br>
 * - Inline styles are used for email client compatibility
 */
export function textToHtml(text: string): string {
  const escaped = escapeHtml(text.trim());

  // Split into paragraphs on 1+ blank lines
  const paragraphs = escaped.split(/\n{2,}/);

  const htmlParas = paragraphs.map((p) => {
    const withBr = p.replace(/\n/g, "<br>");
    return `<p style="margin:0 0 12px;">${withBr}</p>`;
  });

  return `
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;line-height:1.6;color:#111;">
  ${htmlParas.join("\n")}
</div>`.trim();
}

/**
 * Format a timestamp for message display.
 */
export function formatMessageTime(date: Date = new Date()): string {
  return date.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
}
