# 🤖 SharePoint AI Assistant — SPFx Application Customizer

> **"We built a SharePoint-embedded AI assistant that securely retrieves site knowledge and performs contextual actions (like drafting and sending emails) on behalf of users, using delegated Microsoft Graph permissions—frontend-only SPFx and permission-aware by design."**

---

## What Is It?

A **SharePoint Framework (SPFx) Application Customizer** that embeds a context-aware, permission-sensitive AI assistant directly into any modern SharePoint site.

- **No backend required** — 100% frontend, runs in the user's browser
- **Permission-aware** — results are scoped to what the current user can access
- **Three modes**: Site Knowledge, Web Search, and Email Action
- Built with **React**, **TypeScript**, and **Fluent UI v9**

---

## Architecture at a Glance

```
Browser (User Context)
  ├── SPFx ApplicationCustomizer (AssistantApplicationCustomizer.ts)
  │   └── Injects FAB + React Panel into DOM
  ├── React Components (Fluent UI)
  │   ├── AssistantLauncher (FAB + slide-in panel host)
  │   ├── AssistantPanel (orchestrator)
  │   ├── KnowledgeView, EmailDraftView, MessageList, etc.
  ├── Services (all delegated, no secrets)
  │   ├── graphSearchService → /search/query (Graph)
  │   ├── graphUserService  → /me, /me/manager (Graph)
  │   ├── graphMailService  → /me/sendMail (Graph)
  │   ├── spPageService     → DOM + SP REST
  │   └── webSearchService  → Mock (prototype)
  └── State (Zustand, in-memory only, last 5 messages)
```

---

## Permissions

All permissions are **delegated** — the assistant acts as the logged-in user, not as an app with elevated access.

| Permission         | Used For                                           |
|--------------------|----------------------------------------------------|
| `User.Read`        | Read current user's profile (`/me`)                |
| `User.ReadBasic.All` | Read manager info (`/me/manager`)                |
| `Mail.Send`        | Send emails on the user's behalf (`/me/sendMail`)  |
| `Sites.Read.All`   | Search SharePoint site content via Graph           |
| `Files.Read.All`   | Access drive items in search results               |

### Admin Consent Steps

1. Package and deploy the `.sppkg` to the SharePoint App Catalog.
2. In **SharePoint Admin Center** → **Advanced** → **API Access**.
3. Approve each permission request.
4. Users will be prompted for consent on first use if tenant requires it.

---

## Requirements

| Tool     | Version                        |
|----------|--------------------------------|
| Node.js  | **18.x LTS** (v18.17–v18.20)  |
| npm      | 9.x or 10.x                   |
| SPFx     | **1.18.2**                    |
| Gulp CLI | 4.x (`npm i -g gulp-cli`)     |

> ⚠️ **Node version is critical.** SPFx 1.18.2 requires Node 18 LTS.  
> Node 20, 22, or 24 will **not work**. Use `nvm use 18` before installing.

---

## Local Development

### 1. Switch to Node 18 (if using nvm)

```bash
nvm install 18
nvm use 18
node -v   # must show v18.x.x
```

### 2. Install dependencies

```bash
npm install --legacy-peer-deps
```

> `--legacy-peer-deps` is required because Fluent UI v9 has peer dependency declarations  
> that conflict with React 17 in strict npm v9+ resolution.

### 3. Configure your test site

### 2. Configure your test site

Edit `config/serve.json` and replace the `pageUrl` with your tenant:

```json
{
  "serveConfigurations": {
    "default": {
      "pageUrl": "https://yourtenant.sharepoint.com/sites/yoursite/SitePages/Home.aspx"
    }
  }
}
```

### 3. Start local dev server

```bash
gulp serve --nobrowser
```

This starts HTTPS on `localhost:4321`.

### 4. Load the customizer in your browser

Navigate to your SharePoint site and append the debug query string:

```
https://yourtenant.sharepoint.com/sites/yoursite/?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"c5d7e9f1-a2b4-6c8d-0e2f-4a6b8c0d2e4f":{"location":"ClientSideExtension.ApplicationCustomizer"}}
```

The **✦ AI** floating button appears in the bottom-right corner. Click to open!

---

## Building for Production

```bash
# Create production bundle
gulp bundle --ship

# Package .sppkg
gulp package-solution --ship
```

Upload `sharepoint/solution/spfx-ai-assistant.sppkg` to your App Catalog.

---

## Modes Explained

### 🏢 This Site (Knowledge Mode)

Searches your SharePoint site using **Microsoft Graph `/search/query`**, scoped to the current site URL. Results reflect only what the current user can access.

- Click **"Summarize this site"** to get a summary of the current page
- Click **"Key documents"** to find important files
- Type any question about site content

**What you see:** Summary, Key Points, and clickable Citations (webUrls).

### 🌐 Web Mode

Simulated web search (prototype mock). In production, replace `webSearchService.ts` with an APIM-proxied Bing Search or Azure Cognitive Search call.

Results are clearly labeled **🌐 External Source**.

### 📧 Email Action Mode (Hero Feature)

Triggered automatically when you type leave/email phrases.

**Example phrases:**
- "I want to apply leave for tomorrow and day after tomorrow"
- "Request leave for next Friday"
- "I won't be in tomorrow, email my manager"

**Flow:**
1. Intent detected → manager fetched via `/me/manager`
2. Draft email built with correct dates, To, Subject, Body
3. **User must click "Confirm & Send"** — email is never sent automatically
4. On confirm → `POST /me/sendMail` → success animation

---

## Demo Scenarios

### Scenario 1: Knowledge Retrieval
**User types:** "What is the project timeline?"
→ Graph searches the site → Shows summary with citations

### Scenario 2: Restricted Document  
**User asks about a document they can't access:**
→ "No relevant content found in this site. A matching document may exist, but you may not have permission to access it."

### Scenario 3: Page Summary
**User opens a modern SharePoint page and clicks "Summarize this site":**
→ DOM extraction → Key points from page content

### Scenario 4: Leave Email (Hero Demo)
1. Type: "I want to apply leave for tomorrow and day after tomorrow"
2. Assistant fetches manager info
3. Shows draft: To: manager@company.com, Subject: "Leave request for 28 Feb & 01 Mar"
4. Review → Edit if needed → Click "Confirm & Send"
5. Email sent ✅ — green success animation

---

## Security Notes

| Aspect              | How Addressed                                          |
|---------------------|--------------------------------------------------------|
| No backend secrets  | No API keys, no server — purely delegated Graph calls  |
| No auto-send        | Strict user confirmation required before sending email |
| No data caching     | All state is in-memory; cleared on panel close         |
| No app-only perms   | Delegated only — user can only see what they can access|
| External source label| Web results clearly marked "External Source"          |
| No telemetry upload | Console-only logging; no data leaves the browser       |

---

## State Management

Uses **Zustand** (in-memory only). The store holds:
- `mode`: `"site"` | `"web"`
- `messages`: last 5 conversation turns
- `knowledgeSummary`, `emailDraft`, `webResults`
- `loading`, `error`

State is reset when the panel is closed. **Nothing is persisted to localStorage.**

---

## File Structure

```
src/
├── extensions/assistant/
│   ├── AssistantApplicationCustomizer.ts   ← SPFx entry point
│   ├── manifest.json
│   └── components/
│       ├── AssistantLauncher.tsx            ← FAB + panel host
│       ├── AssistantPanel.tsx               ← Main orchestrator
│       ├── Header.tsx                       ← Site name + mode toggle
│       ├── ModeToggle.tsx                   ← Site/Web toggle
│       ├── Suggestions.tsx                  ← Smart suggestion buttons
│       ├── KnowledgeView.tsx                ← Summary + citations
│       ├── EmailDraftView.tsx               ← Email draft + send
│       ├── MessageList.tsx                  ← Chat history
│       ├── Loading.tsx                      ← Spinner
│       └── styles.module.scss              ← All styles
├── services/
│   ├── graphClient.ts                       ← MSGraphClientFactory wrapper
│   ├── graphSearchService.ts                ← /search/query
│   ├── graphUserService.ts                  ← /me, /me/manager
│   ├── graphMailService.ts                  ← /me/sendMail
│   ├── spPageService.ts                     ← DOM + SP REST
│   ├── webSearchService.ts                  ← Mock web search
│   ├── ai/
│   │   ├── intentDetector.ts               ← Regex intent parsing
│   │   ├── summarizer.ts                   ← Extractive summarization
│   │   └── promptBuilder.ts               ← Future LLM prompt composer
│   └── types.ts                            ← Shared TypeScript types
├── state/
│   ├── store.ts                            ← Zustand store
│   └── models.ts                           ← State interface
└── utils/
    ├── formatting.ts                        ← Date resolution
    ├── permissions.ts                       ← Error message helpers
    └── telemetry.ts                         ← Console telemetry
```

---

## Future Roadmap

- [ ] **Azure OpenAI integration** — APIM-proxied LLM for higher-quality summaries (no secrets in client; token issued by managed identity)
- [ ] **Teams integration** — Surface assistant in Teams tab or message extension
- [ ] **Multi-site mode** — Search across multiple sites with cross-site scoping
- [ ] **Planner / Tasks** — Create tasks from AI-detected action items
- [ ] **Real web search** — Replace mock with Bing Search API via APIM proxy
- [ ] **Voice input** — Web Speech API for hands-free queries
- [ ] **Adaptive Cards** — Rich email previews with Teams-style cards

---

## Troubleshooting

| Issue | Fix |
|-------|-----|
| `No matching version found for @microsoft/eslint-config-spfx` | Use `npm install --legacy-peer-deps` and ensure you have the fixed `package.json` (SPFx 1.18.2) |
| `npm install` fails with peer dep errors | Always run `npm install --legacy-peer-deps` for SPFx projects |
| Node version error during install | Run `nvm use 18` — SPFx 1.18.2 requires Node 18 LTS exactly |
| FAB doesn't appear | Check the `customActions` query string matches the manifest `id` |
| Graph 401 errors | Ensure admin has approved permissions in SharePoint Admin → API Access |
| Graph 403 errors | User lacks `Sites.Read.All` or file-level permissions |
| Manager returns null | User has no manager configured in Azure AD |
| Search returns 0 results | Site must be indexed; wait 15–30 mins after content creation |
| Email fails to send | `Mail.Send` permission must be granted; check admin consent |
| DOM extraction fails | Modern pages only; classic pages not supported |
| `gulp serve` cert error | Trust the dev cert: `gulp trust-dev-cert` |

---

*Built with ❤️ using SPFx 1.19, React 17, Fluent UI v9, Zustand, and Microsoft Graph.*
