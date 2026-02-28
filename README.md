# 🚀 SharePoint Smart Assistant (AI-Powered Enterprise Knowledge Copilot)

## 👥 Team Details

- **Team Name:** Team Nexus(Tech-Leak Shield)  
- **Members:**  
  - Sriman
  - Sanju
  - Geethika
  - Mounika   
- **Domain Category:** Enterprise AI / Multi-Agent Systems
- **Demo Video:** [Demo Video Link – To Be Updated]

---

## 🎯 Problem Statement

SharePoint is the centralized knowledge hub for organizations. It stores:

- Policies
- SOPs
- Project documents  
- Reports
- Contracts
- Internal communication  

While security and permissions are fully managed by Microsoft 365, users still face challenges:
- Difficulty searching for documents
- Time-consuming folder navigation
- Large documents requiring manual reading
- Manual drafting of emails based on SharePoint content

The process is repetitive, inefficient, and slows decision-making.
---

## 💡 Solution Overview

We built a **SharePoint Smart Assistant** that transforms SharePoint into an intelligent conversational system.
The assistant:
  
  1. Understands natural language queries  
  2. Searches across SharePoint based on user permissions
  3. Summarizes large documents instantly 
  4. Extracts relevant sections from files
  5. Drafts professional emails on behalf of the user
  6. Maintains enterprise-grade security compliance

The system respects Microsoft 365 permissions — if a user does not have access to a document, the assistant cannot retrieve it.

---

## 🏗 Architecture

📁 Architecture Diagram: `/architecture/architecture.png`

### Core Components

- SPFx React Frontend (Floating Assistant UI) 
- Azure OpenAI Integration
- Microsoft Graph API
- SharePoint REST API
- Orchestration Layer
- Context Memory Layer
- Role-Based Permission Validation

### Application Flow

1. User asks a natural language question
2. Assistant authenticates via Microsoft 365
3. Query is analyzed using Azure OpenAI
4. Microsoft Graph fetches permitted SharePoint data
5. Relevant documents are retrieved
6. AI summarizes or generates requested output
7. User receives contextual response

---

## 🛠 Tech Stack

| Layer | Technology |
|-------|------------|
| Frontend | 	React (SPFx) |
| Backend Logic | TypeScript |
| AI Model | Azure OpenAI GPT-4 |
| APIs | Microsoft Graph API |
| Platform | SharePoint Online |
| Authentication | Azure AD |
| Deployment | SharePoint App Catalog |

---

## 📂 Project Structure

```
spfx-ai-assistant/
│
├── README.md
├── package.json
├── .env.example
│
├── src/
│   ├── extensions/
│   │   └── aiAssistant/
│   │       ├── AiAssistantApplicationCustomizer.ts
│   │       ├── components/
│   │       │   ├── ChatBox.tsx
│   │       │   ├── Message.tsx
│   │       │   └── InputBox.tsx
│   │       ├── services/
│   │       │   ├── GraphService.ts
│   │       │   └── OpenAIService.ts
│   │       └── utils/
│   │           └── config.ts
│
├── architecture/
│   └── architecture.png
```

---

### 🚨 Mandatory Files for All Submissions

The following files **must be present** in every submission:

- `README.md`
- `requirements.txt` (or `package.json` for Node projects)
- `.env.example`
-  Clear entry point inside `src/`
  
All other folders (e.g., `data/`, `tests/`, `notebooks/`, etc.) may vary depending on the project.
**Submissions missing mandatory files may not be evaluated.**

---

## ⚙️ Setup Instructions

## 1️ Verify Required Software

- Node.js (v18 or above)
- SharePoint Online tenant
- SPFx environment setup
- Azure OpenAI resource
- Azure AD App Registration

### 1️⃣ Clone Repository

```bash
git clone https://github.com/K-Sriman/spfx-ai-assistant.git
cd spfx-ai-assistant
```

### 2️⃣ Install Dependencies

```bash
npm install
```

### 3️⃣ Configure Environment Variables

Create `.env` file from `.env.example`

Example:

```
AZURE_OPENAI_KEY=your_key_here
```

---

## ▶️ Entry Point

Run the application:

```bash
gulp serve
```

SPFx Application Customizer Extension:

```
AiAssistantApplicationCustomizer.ts
```
The assistant appears as a floating AI chat bubble on all SharePoint pages.

---

## 🧪 How to Test

### Example Queries
Example Queries
  - Give me the latest finance policy.
  - Summarize the Q4 sales report.
  - Draft an email requesting budget approval.
  - Show me HR leave policy applicable to my role.
  - Extract key points from the onboarding document.
    
## ⚠️ Known Limitations
- Requires OpenAI API access  
- layout is not pixel-perfect  
- Responses may occasionally be inaccurate due to token limitations or model constraints
- Performance depends on model availability and API limits

---

## 🔮 Future Improvements

- Multi-agent workflow automation  
- Automated approval flow generation  
- Context-based role intelligence 
- SharePoint list data analytics
- Power Automate integration
- Reinforcement feedback learning  

---

🌟 Impact
This solution converts SharePoint from a static document repository into an intelligent enterprise copilot.
Instead of training users to search better,
we make SharePoint understand users better.

---
