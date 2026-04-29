# MCP Office Extension

A Microsoft Office add-in that brings AI chat, selection actions, and MCP (Model Context Protocol) tool support directly into Word, Excel, and PowerPoint. Bring your own API key — no backend required.

---

## Features

- **Chat** — Streaming AI chat with insert-to-document support
- **Actions** — One-click actions on selected text: Summarize, Rewrite, Fix Grammar, Translate, Expand, Shorten, Improve
- **MCP Servers** — Connect any remote MCP server and give the AI access to external tools
- **Settings** — Configure your AI provider, API key, model, and system prompt

## Supported AI Providers

| Provider | Notes |
|---|---|
| OpenRouter | Recommended — access Claude, GPT-4, Gemini and more with one key |
| OpenAI | Direct GPT-4o, o1, etc. |
| Google Gemini | Direct Gemini 2.0 Flash, 1.5 Pro, etc. |
| OpenAI-compatible | Ollama, Groq, Together, any custom endpoint |

---

## Installation (End Users)

No npm or Node.js needed. Just sideload the manifest file into Office.

### Step 1 — Download the manifest

Download `manifest.xml` from this repository.

Open it in a text editor and replace both occurrences of `https://localhost:3000` with the hosted URL where the add-in is deployed (ask whoever shared this with you for the URL).

### Step 2 — Sideload into Office

**Word / Excel / PowerPoint on Windows or Mac:**

1. Open the Office app
2. Go to **Insert → Add-ins → My Add-ins**
3. Click **Upload My Add-in**
4. Select the `manifest.xml` file you downloaded
5. Click **Upload**

The **MCP AI** button will appear in the **Home** ribbon. Click it to open the side panel.

**Word Online / Excel Online:**

1. Go to **Insert → Add-ins**
2. Click **Upload My Add-in**
3. Select `manifest.xml`

### Step 3 — Configure your API key

1. Click **MCP AI** in the ribbon to open the panel
2. Go to the **Settings** tab
3. Select your provider (OpenRouter is recommended)
4. Enter your API key
5. Choose a model
6. Click **Save**

> **Getting an OpenRouter key:** Sign up at [openrouter.ai](https://openrouter.ai), go to Keys, and create a new key. OpenRouter gives access to Claude, GPT-4, Gemini and many other models under one key.

---

## Development Setup

Requirements: Node.js 18+ and npm.

### 1. Clone and install

```bash
git clone https://github.com/askbard/mcpwordextension.git
cd mcpwordextension
npm install
```

### 2. Start the dev server

```bash
npm run dev
```

This starts a local HTTPS server at `https://localhost:3000`. The HTTPS certificate is generated automatically by `vite-plugin-mkcert` on first run — accept the certificate prompt if your browser asks.

> The dev server must keep running while you test in Office. Stop it with `Ctrl+C` when done.

### 3. Sideload the manifest

The default `manifest.xml` already points to `https://localhost:3000`, so no edits needed for local development.

Follow the sideload steps in the Installation section above using the `manifest.xml` from the repo root.

### 4. Make changes

Edit files under `src/`. Vite hot-reloads automatically — just refresh the add-in panel in Office to see changes.

---

## Deploying to Production

### Build

```bash
npm run build
```

Output goes to the `dist/` folder.

### Host the files

Upload the contents of `dist/` to any static HTTPS host:

- **Netlify** — drag and drop the `dist/` folder at [app.netlify.com/drop](https://app.netlify.com/drop)
- **GitHub Pages** — push `dist/` to the `gh-pages` branch
- **Vercel** — `npx vercel` in the project root

### Update the manifest

Open `manifest.xml` and replace all instances of `https://localhost:3000` with your deployed URL, for example:

```xml
<!-- Before -->
DefaultValue="https://localhost:3000/src/taskpane/index.html"

<!-- After -->
DefaultValue="https://your-app.netlify.app/src/taskpane/index.html"
```

Share the updated `manifest.xml` with your users.

---

## Adding MCP Servers

1. Open the panel and go to the **MCP** tab
2. Click **Add**
3. Enter a name and the server URL (must be an HTTP/SSE endpoint, e.g. `https://your-mcp-server.com/mcp`)
4. Optionally add auth headers (e.g. `{"Authorization": "Bearer your-token"}`)
5. Click **Add Server**, then click the sync icon to connect
6. Available tools appear listed under the server once connected

Connected tools are automatically available to the AI in the **Chat** tab.

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
manifest.xml                    # Office add-in manifest
```

---

## Security Notes

- API keys are stored in your browser's `localStorage` — they never leave your machine
- MCP servers are called directly from the browser; ensure your MCP server supports CORS if hosted remotely
- The add-in requests `ReadWriteDocument` permission to read selections and insert text
