# MCP Office Extension

Use any AI model with MCP tool support directly inside Microsoft Word, Excel, and PowerPoint. Bring your own API key — no backend, no subscription required.

**Live add-in:** `https://askbard.github.io/mcpwordextension/`

---

## Features

- **Chat** — Streaming AI chat with insert-to-document support
- **Actions** — One-click actions on selected text: Summarize, Rewrite, Fix Grammar, Translate, Expand, Shorten, Improve
- **MCP Servers** — Connect any remote MCP server and give the AI access to external tools
- **Settings** — Configure your AI provider, API key, model, and system prompt

---

## Quick Install (End Users)

> No Node.js or technical setup needed. Just two steps.

### Step 1 — Download the manifest

👉 [Download manifest-production.xml](https://raw.githubusercontent.com/askbard/mcpwordextension/main/manifest-production.xml)

Right-click the link → **Save link as** → save as `manifest-production.xml`

### Step 2 — Sideload into Office

**Windows (Desktop)**

1. Open Word, Excel, or PowerPoint
2. Click **Insert** in the top menu
3. Click **Add-ins** → **My Add-ins**
4. Click **Upload My Add-in**
5. Select the `manifest-production.xml` file you just downloaded
6. Click **Upload**

The **MCP AI** button will appear in the **Home** ribbon. Click it to open the panel.

**Mac (Desktop) — Method 1: Upload My Add-in**

1. Open Word, Excel, or PowerPoint
2. Click **Insert** in the menu bar
3. Click **Add-ins**
4. In the dialog, click **"..."** or look for **Upload My Add-in**
5. Select `manifest-production.xml`

**Mac (Desktop) — Method 2: Manual folder (most reliable)**

If Method 1 doesn't work:

1. Quit Word (or Excel/PowerPoint) completely
2. Open **Finder** → press **Cmd+Shift+G** → paste the path for your app:
   - Word: `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`
   - Excel: `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
   - PowerPoint: `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`
3. If the `wef` folder doesn't exist, create it
4. Copy `manifest-production.xml` into that folder
5. Reopen the Office app — go to **Insert → Add-ins → My Add-ins** to find it

**Word Online / Excel Online / PowerPoint Online**

1. Click **Insert** in the top menu
2. Click **Add-ins**
3. Click **Upload My Add-in**
4. Select `manifest-production.xml`
5. Click **Upload**

---

### Step 3 — Get an API key

OpenRouter is recommended — one key gives access to Claude, GPT-4, Gemini, and 100+ other models.

1. Go to [openrouter.ai](https://openrouter.ai) and sign up (free)
2. Go to **Keys** → **Create Key**
3. Copy the key (starts with `sk-or-...`)

### Step 4 — Configure the add-in

1. In Office, click **MCP AI** in the ribbon to open the panel
2. Go to the **Settings** tab
3. Select **OpenRouter** as the provider
4. Paste your API key
5. Choose a model (e.g. `anthropic/claude-3.5-sonnet`)
6. Click **Save**

You're ready to go. Start chatting in the **Chat** tab or select text and use **Actions**.

---

## Using MCP Servers (Optional)

MCP servers give the AI access to external tools like web search, databases, or your own APIs.

1. Open the panel and go to the **MCP** tab
2. Click **Add**
3. Enter a name and the server URL (e.g. `https://your-mcp-server.com/mcp`)
4. Optionally add auth headers (e.g. `{"Authorization": "Bearer your-token"}`)
5. Click **Add Server**, then click the **sync icon** to connect
6. Connected tools are automatically available to the AI in the Chat tab

---

## Supported AI Providers

| Provider | Notes |
|---|---|
| OpenRouter | Recommended — Claude, GPT-4, Gemini and 100+ models with one key |
| OpenAI | Direct GPT-4o, o1, etc. |
| Google Gemini | Direct Gemini 2.0 Flash, 1.5 Pro, etc. |
| OpenAI-compatible | Ollama, Groq, Together, any custom endpoint |

---

## Developer Setup

Requirements: Node.js 18+ and npm.

### Run locally

```bash
git clone https://github.com/askbard/mcpwordextension.git
cd mcpwordextension
npm install
npm run dev
```

This starts a local HTTPS server at `https://localhost:3000`. Sideload `manifest.xml` (not the production one) for local development.

> The dev server must keep running while you test in Office. Stop it with `Ctrl+C` when done.

### Build for production

```bash
npm run build
```

Output goes to `dist/`. Deploy to any static HTTPS host and update the URLs in `manifest-production.xml`.

---

## Project Structure

```
src/
├── providers/index.ts          # AI streaming via OpenAI-compatible SDK
├── mcp/client.ts               # MCP HTTP+SSE client
├── office/bridge.ts            # Office.js helpers for Word, Excel, PowerPoint
├── storage/settings.ts         # localStorage for settings and MCP configs
└── taskpane/
    ├── App.tsx                 # 4-tab shell
    ├── hooks/useAgentLoop.ts   # Agentic tool-calling loop
    └── components/
        ├── ChatPanel.tsx
        ├── ActionsPanel.tsx
        ├── McpPanel.tsx
        └── SettingsPanel.tsx
manifest.xml                    # Manifest for local development
manifest-production.xml         # Manifest for GitHub Pages deployment
```

---

## Security Notes

- API keys are stored in your browser's `localStorage` — they never leave your machine
- MCP servers are called directly from the browser; ensure your MCP server supports CORS
- The add-in requests `ReadWriteDocument` permission to read selections and insert text
