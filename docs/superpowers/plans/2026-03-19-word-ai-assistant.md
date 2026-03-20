# Word AI Assistant Add-in Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build an Office.js Word Add-in with a chat-based TaskPane that reads document structure, sends it with user-provided raw data to an OpenAI-compatible API, and writes AI responses back into the document.

**Architecture:** Office.js Web Add-in using React+TypeScript, bundled with Webpack, served via webpack-dev-server for development. The add-in runs in a TaskPane (side panel) and communicates with an on-prem OpenAI-compatible API via fetch. Document manipulation uses the Office.js Word API.

**Tech Stack:** TypeScript, React 18, Office.js (office-js), Webpack 5, CSS

---

## File Map

| File | Responsibility |
|------|---------------|
| `manifest.xml` | Office Add-in manifest — declares TaskPane, icons, permissions |
| `package.json` | Dependencies and scripts |
| `tsconfig.json` | TypeScript config |
| `webpack.config.js` | Webpack bundler config with HtmlWebpackPlugin, dev-server HTTPS |
| `src/taskpane/index.html` | HTML shell for TaskPane |
| `src/taskpane/index.tsx` | React mount point, Office.onReady |
| `src/taskpane/App.tsx` | Main app with tab navigation (Chat / Settings) |
| `src/taskpane/components/ChatPanel.tsx` | Chat UI: message list, input box, send button |
| `src/taskpane/components/MessageBubble.tsx` | Individual message display (user/assistant) |
| `src/taskpane/components/SettingsPanel.tsx` | AI server URL and model name configuration |
| `src/taskpane/services/settings.ts` | localStorage read/write for settings |
| `src/taskpane/services/aiClient.ts` | OpenAI-compatible chat completions API caller |
| `src/taskpane/services/wordDocument.ts` | Office.js Word API: read doc structure, write content |
| `src/taskpane/styles/app.css` | All styles |
| `assets/icon-16.png` | 16x16 add-in icon |
| `assets/icon-32.png` | 32x32 add-in icon |
| `assets/icon-80.png` | 80x80 add-in icon |

---

## Chunk 1: Project Scaffolding & Dev Server

### Task 1: Initialize project and install dependencies

**Files:**
- Create: `package.json`
- Create: `tsconfig.json`
- Create: `webpack.config.js`

- [ ] **Step 1: Initialize npm project**

```bash
cd /Users/junyung.ahn/Desktop/Work/09_word_extension
npm init -y
```

- [ ] **Step 2: Install dependencies**

```bash
npm install react react-dom office-addin-mock
npm install -D typescript @types/react @types/react-dom \
  webpack webpack-cli webpack-dev-server \
  html-webpack-plugin copy-webpack-plugin \
  ts-loader css-loader style-loader \
  @types/office-js
```

- [ ] **Step 3: Create tsconfig.json**

```json
{
  "compilerOptions": {
    "target": "ES2020",
    "module": "ES2020",
    "moduleResolution": "bundler",
    "jsx": "react-jsx",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "outDir": "./dist",
    "rootDir": "./src",
    "lib": ["ES2020", "DOM", "DOM.Iterable"]
  },
  "include": ["src/**/*"],
  "exclude": ["node_modules", "dist"]
}
```

- [ ] **Step 4: Create webpack.config.js**

```js
const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = {
  mode: "development",
  entry: {
    taskpane: "./src/taskpane/index.tsx",
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].bundle.js",
    clean: true,
  },
  resolve: {
    extensions: [".ts", ".tsx", ".js", ".jsx"],
  },
  module: {
    rules: [
      { test: /\.tsx?$/, use: "ts-loader", exclude: /node_modules/ },
      { test: /\.css$/, use: ["style-loader", "css-loader"] },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: "./src/taskpane/index.html",
      filename: "taskpane.html",
      chunks: ["taskpane"],
    }),
    new CopyWebpackPlugin({
      patterns: [{ from: "assets", to: "assets" }],
    }),
  ],
  devServer: {
    static: path.resolve(__dirname, "dist"),
    port: 3000,
    https: true,
    headers: { "Access-Control-Allow-Origin": "*" },
  },
  devtool: "source-map",
};
```

- [ ] **Step 5: Create placeholder icons**

Create `assets/icon-16.png`, `assets/icon-32.png`, `assets/icon-80.png` (simple placeholder PNGs).

- [ ] **Step 6: Commit**

```bash
git init
git add package.json tsconfig.json webpack.config.js assets/
git commit -m "chore: initialize project with typescript, webpack, react deps"
```

---

### Task 2: Create manifest.xml and HTML shell

**Files:**
- Create: `manifest.xml`
- Create: `src/taskpane/index.html`

- [ ] **Step 1: Create manifest.xml**

```xml
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xsi:type="TaskPaneApp">
  <Id>a1b2c3d4-e5f6-7890-abcd-ef1234567890</Id>
  <Version>1.0.0</Version>
  <ProviderName>Internal</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="AI Assistant"/>
  <Description DefaultValue="AI-powered document writing assistant"/>
  <IconUrl DefaultValue="https://localhost:3000/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://localhost:3000/assets/icon-80.png"/>
  <SupportUrl DefaultValue="https://localhost:3000"/>
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

- [ ] **Step 2: Create index.html**

```html
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>AI Assistant</title>
  <script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
</head>
<body>
  <div id="root"></div>
</body>
</html>
```

- [ ] **Step 3: Commit**

```bash
git add manifest.xml src/taskpane/index.html
git commit -m "chore: add Office Add-in manifest and HTML shell"
```

---

### Task 3: Create React mount point and minimal App

**Files:**
- Create: `src/taskpane/index.tsx`
- Create: `src/taskpane/App.tsx`
- Create: `src/taskpane/styles/app.css`

- [ ] **Step 1: Create index.tsx**

```tsx
import React from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import "./styles/app.css";

Office.onReady(() => {
  const root = createRoot(document.getElementById("root")!);
  root.render(<App />);
});
```

- [ ] **Step 2: Create minimal App.tsx**

```tsx
import React, { useState } from "react";

type Tab = "chat" | "settings";

export function App() {
  const [activeTab, setActiveTab] = useState<Tab>("chat");

  return (
    <div className="app">
      <header className="app-header">
        <button
          className={`tab-btn ${activeTab === "chat" ? "active" : ""}`}
          onClick={() => setActiveTab("chat")}
        >
          Chat
        </button>
        <button
          className={`tab-btn ${activeTab === "settings" ? "active" : ""}`}
          onClick={() => setActiveTab("settings")}
        >
          Settings
        </button>
      </header>
      <main className="app-content">
        {activeTab === "chat" ? (
          <div>Chat panel placeholder</div>
        ) : (
          <div>Settings panel placeholder</div>
        )}
      </main>
    </div>
  );
}
```

- [ ] **Step 3: Create app.css with basic styling**

```css
* { margin: 0; padding: 0; box-sizing: border-box; }

body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  font-size: 14px;
  color: #333;
  background: #fff;
}

.app {
  display: flex;
  flex-direction: column;
  height: 100vh;
}

.app-header {
  display: flex;
  border-bottom: 1px solid #e0e0e0;
  background: #f5f5f5;
}

.tab-btn {
  flex: 1;
  padding: 10px;
  border: none;
  background: transparent;
  cursor: pointer;
  font-size: 14px;
  color: #666;
}

.tab-btn.active {
  color: #0078d4;
  border-bottom: 2px solid #0078d4;
  font-weight: 600;
}

.app-content {
  flex: 1;
  overflow-y: auto;
  padding: 12px;
}
```

- [ ] **Step 4: Build and verify dev server starts**

Run: `npx webpack serve`
Expected: Dev server starts on https://localhost:3000, `taskpane.html` loads with tab UI.

- [ ] **Step 5: Verify sideload in Mac Word**

1. Open Word on Mac
2. Insert > Add-ins > My Add-ins > Upload My Add-in
3. Select `manifest.xml`
4. TaskPane should open with the tab UI

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/index.tsx src/taskpane/App.tsx src/taskpane/styles/app.css
git commit -m "feat: add React app shell with tab navigation"
```

---

## Chunk 2: Settings Service & Panel

### Task 4: Implement settings service

**Files:**
- Create: `src/taskpane/services/settings.ts`

- [ ] **Step 1: Create settings.ts**

```ts
export interface Settings {
  apiUrl: string;
  modelName: string;
}

const STORAGE_KEY = "word-ai-assistant-settings";

const DEFAULT_SETTINGS: Settings = {
  apiUrl: "http://localhost:11434/v1/chat/completions",
  modelName: "llama3",
};

export function loadSettings(): Settings {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (stored) {
      return { ...DEFAULT_SETTINGS, ...JSON.parse(stored) };
    }
  } catch {
    // ignore parse errors
  }
  return { ...DEFAULT_SETTINGS };
}

export function saveSettings(settings: Settings): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}
```

- [ ] **Step 2: Commit**

```bash
git add src/taskpane/services/settings.ts
git commit -m "feat: add settings service with localStorage persistence"
```

---

### Task 5: Implement SettingsPanel component

**Files:**
- Create: `src/taskpane/components/SettingsPanel.tsx`
- Modify: `src/taskpane/App.tsx`

- [ ] **Step 1: Create SettingsPanel.tsx**

```tsx
import React, { useState, useEffect } from "react";
import { loadSettings, saveSettings, Settings } from "../services/settings";

export function SettingsPanel() {
  const [settings, setSettings] = useState<Settings>(loadSettings);
  const [saved, setSaved] = useState(false);

  const handleSave = () => {
    saveSettings(settings);
    setSaved(true);
    setTimeout(() => setSaved(false), 2000);
  };

  return (
    <div className="settings-panel">
      <div className="setting-group">
        <label className="setting-label">AI Server URL</label>
        <input
          className="setting-input"
          type="text"
          value={settings.apiUrl}
          onChange={(e) => setSettings({ ...settings, apiUrl: e.target.value })}
          placeholder="http://localhost:11434/v1/chat/completions"
        />
      </div>
      <div className="setting-group">
        <label className="setting-label">Model Name</label>
        <input
          className="setting-input"
          type="text"
          value={settings.modelName}
          onChange={(e) => setSettings({ ...settings, modelName: e.target.value })}
          placeholder="llama3"
        />
      </div>
      <button className="save-btn" onClick={handleSave}>
        {saved ? "Saved!" : "Save Settings"}
      </button>
    </div>
  );
}
```

- [ ] **Step 2: Add settings styles to app.css**

Append to `src/taskpane/styles/app.css`:

```css
/* Settings */
.settings-panel {
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.setting-group {
  display: flex;
  flex-direction: column;
  gap: 4px;
}

.setting-label {
  font-size: 12px;
  font-weight: 600;
  color: #555;
}

.setting-input {
  padding: 8px;
  border: 1px solid #ccc;
  border-radius: 4px;
  font-size: 13px;
}

.setting-input:focus {
  outline: none;
  border-color: #0078d4;
}

.save-btn {
  padding: 8px 16px;
  background: #0078d4;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 14px;
}

.save-btn:hover {
  background: #006abc;
}
```

- [ ] **Step 3: Wire SettingsPanel into App.tsx**

Replace the settings placeholder in `App.tsx`:

```tsx
import { SettingsPanel } from "./components/SettingsPanel";

// In the render, replace <div>Settings panel placeholder</div> with:
<SettingsPanel />
```

- [ ] **Step 4: Verify in browser**

Run: `npx webpack serve`
Navigate to https://localhost:3000/taskpane.html, click Settings tab, enter URL and model, click Save. Refresh page — settings should persist.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/components/SettingsPanel.tsx src/taskpane/styles/app.css src/taskpane/App.tsx
git commit -m "feat: add settings panel with API URL and model config"
```

---

## Chunk 3: Chat UI

### Task 6: Implement MessageBubble component

**Files:**
- Create: `src/taskpane/components/MessageBubble.tsx`

- [ ] **Step 1: Create MessageBubble.tsx**

```tsx
import React from "react";

export interface Message {
  role: "user" | "assistant";
  content: string;
}

export function MessageBubble({ role, content }: Message) {
  return (
    <div className={`message ${role}`}>
      <div className="message-role">{role === "user" ? "You" : "AI"}</div>
      <div className="message-content">{content}</div>
    </div>
  );
}
```

- [ ] **Step 2: Add message styles to app.css**

```css
/* Messages */
.message {
  margin-bottom: 12px;
  padding: 8px 12px;
  border-radius: 8px;
  max-width: 95%;
}

.message.user {
  background: #e8f0fe;
  align-self: flex-end;
}

.message.assistant {
  background: #f0f0f0;
  align-self: flex-start;
}

.message-role {
  font-size: 11px;
  font-weight: 600;
  color: #888;
  margin-bottom: 4px;
}

.message-content {
  white-space: pre-wrap;
  word-break: break-word;
  line-height: 1.5;
}
```

- [ ] **Step 3: Commit**

```bash
git add src/taskpane/components/MessageBubble.tsx src/taskpane/styles/app.css
git commit -m "feat: add MessageBubble component"
```

---

### Task 7: Implement ChatPanel component

**Files:**
- Create: `src/taskpane/components/ChatPanel.tsx`
- Modify: `src/taskpane/App.tsx`

- [ ] **Step 1: Create ChatPanel.tsx**

```tsx
import React, { useState, useRef, useEffect } from "react";
import { Message, MessageBubble } from "./MessageBubble";

interface ChatPanelProps {
  onSend: (userMessage: string) => Promise<string>;
}

export function ChatPanel({ onSend }: ChatPanelProps) {
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  const handleSend = async () => {
    const text = input.trim();
    if (!text || loading) return;

    setInput("");
    setMessages((prev) => [...prev, { role: "user", content: text }]);
    setLoading(true);

    try {
      const reply = await onSend(text);
      setMessages((prev) => [...prev, { role: "assistant", content: reply }]);
    } catch (err: any) {
      setMessages((prev) => [
        ...prev,
        { role: "assistant", content: `Error: ${err.message}` },
      ]);
    } finally {
      setLoading(false);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  return (
    <div className="chat-panel">
      <div className="messages-list">
        {messages.map((msg, i) => (
          <MessageBubble key={i} {...msg} />
        ))}
        {loading && (
          <div className="message assistant">
            <div className="message-role">AI</div>
            <div className="message-content loading-dots">Thinking...</div>
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>
      <div className="chat-input-area">
        <textarea
          className="chat-input"
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={handleKeyDown}
          placeholder="Paste your raw data here..."
          rows={3}
        />
        <button
          className="send-btn"
          onClick={handleSend}
          disabled={loading || !input.trim()}
        >
          Send
        </button>
      </div>
    </div>
  );
}
```

- [ ] **Step 2: Add chat styles to app.css**

```css
/* Chat */
.chat-panel {
  display: flex;
  flex-direction: column;
  height: 100%;
}

.messages-list {
  flex: 1;
  overflow-y: auto;
  display: flex;
  flex-direction: column;
  padding-bottom: 8px;
}

.chat-input-area {
  display: flex;
  gap: 8px;
  padding-top: 8px;
  border-top: 1px solid #e0e0e0;
}

.chat-input {
  flex: 1;
  padding: 8px;
  border: 1px solid #ccc;
  border-radius: 4px;
  font-size: 13px;
  font-family: inherit;
  resize: none;
}

.chat-input:focus {
  outline: none;
  border-color: #0078d4;
}

.send-btn {
  padding: 8px 16px;
  background: #0078d4;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 14px;
  align-self: flex-end;
}

.send-btn:disabled {
  background: #ccc;
  cursor: not-allowed;
}

.loading-dots {
  color: #888;
  font-style: italic;
}
```

- [ ] **Step 3: Wire ChatPanel into App.tsx with a mock onSend**

Update `App.tsx` to import `ChatPanel` and pass a temporary mock handler:

```tsx
import { ChatPanel } from "./components/ChatPanel";

// Replace chat placeholder with:
<ChatPanel onSend={async (msg) => `Echo: ${msg}`} />
```

- [ ] **Step 4: Verify in browser**

Run dev server, type a message, press Enter. Should see user bubble and "Echo: ..." reply.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/components/ChatPanel.tsx src/taskpane/styles/app.css src/taskpane/App.tsx
git commit -m "feat: add ChatPanel with message history and input"
```

---

## Chunk 4: AI Client Service

### Task 8: Implement aiClient service

**Files:**
- Create: `src/taskpane/services/aiClient.ts`

- [ ] **Step 1: Create aiClient.ts**

```ts
import { loadSettings } from "./settings";

interface ChatMessage {
  role: "system" | "user" | "assistant";
  content: string;
}

interface ChatCompletionResponse {
  choices: Array<{
    message: {
      role: string;
      content: string;
    };
  }>;
}

export async function sendChatRequest(
  systemPrompt: string,
  userMessage: string
): Promise<string> {
  const { apiUrl, modelName } = loadSettings();

  if (!apiUrl) {
    throw new Error("AI Server URL is not configured. Go to Settings tab.");
  }

  const messages: ChatMessage[] = [
    { role: "system", content: systemPrompt },
    { role: "user", content: userMessage },
  ];

  const response = await fetch(apiUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: modelName,
      messages,
      temperature: 0.3,
    }),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`API error (${response.status}): ${errorText}`);
  }

  const data: ChatCompletionResponse = await response.json();

  if (!data.choices || data.choices.length === 0) {
    throw new Error("No response from AI model");
  }

  return data.choices[0].message.content;
}
```

- [ ] **Step 2: Commit**

```bash
git add src/taskpane/services/aiClient.ts
git commit -m "feat: add OpenAI-compatible API client service"
```

---

## Chunk 5: Word Document Service & Integration

### Task 9: Implement wordDocument service

**Files:**
- Create: `src/taskpane/services/wordDocument.ts`

- [ ] **Step 1: Create wordDocument.ts**

```ts
export async function readDocumentStructure(): Promise<string> {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");

    const tables = body.tables;
    tables.load("count");

    await context.sync();

    const parts: string[] = [];
    parts.push("=== DOCUMENT TEXT ===");
    parts.push(body.text);

    if (tables.count > 0) {
      for (let i = 0; i < tables.count; i++) {
        const table = tables.items[i];
        table.load("rowCount");
        table.load("values");
      }
      await context.sync();

      for (let i = 0; i < tables.count; i++) {
        const table = tables.items[i];
        parts.push(`\n=== TABLE ${i + 1} ===`);
        const values = table.values;
        for (let r = 0; r < values.length; r++) {
          parts.push(values[r].join(" | "));
        }
      }
    }

    return parts.join("\n");
  });
}

export async function applyAIResponse(response: string): Promise<void> {
  return Word.run(async (context) => {
    const body = context.document.body;
    const tables = body.tables;
    tables.load("count");
    await context.sync();

    // Try to parse structured JSON response
    const jsonMatch = response.match(/```json\s*([\s\S]*?)\s*```/);
    if (jsonMatch) {
      const instructions = JSON.parse(jsonMatch[1]);
      await applyInstructions(context, instructions, tables);
    } else {
      // Fallback: insert response as text at end of document
      body.insertParagraph(response, Word.InsertLocation.end);
    }

    await context.sync();
  });
}

interface WriteInstruction {
  type: "table_cell" | "paragraph" | "replace";
  tableIndex?: number;
  row?: number;
  col?: number;
  paragraphIndex?: number;
  searchText?: string;
  value: string;
}

async function applyInstructions(
  context: Word.RequestContext,
  instructions: WriteInstruction[],
  tables: Word.TableCollection
): Promise<void> {
  for (const inst of instructions) {
    if (inst.type === "table_cell" && inst.tableIndex !== undefined) {
      const table = tables.items[inst.tableIndex];
      const cell = table.getCell(inst.row!, inst.col!);
      cell.body.clear();
      cell.body.insertText(inst.value, Word.InsertLocation.start);
    } else if (inst.type === "replace" && inst.searchText) {
      const results = context.document.body.search(inst.searchText, {
        matchCase: false,
        matchWholeWord: false,
      });
      results.load("items");
      await context.sync();
      if (results.items.length > 0) {
        results.items[0].insertText(inst.value, Word.InsertLocation.replace);
      }
    } else if (inst.type === "paragraph") {
      const body = context.document.body;
      body.insertParagraph(inst.value, Word.InsertLocation.end);
    }
  }
}
```

- [ ] **Step 2: Commit**

```bash
git add src/taskpane/services/wordDocument.ts
git commit -m "feat: add Word document read/write service with Office.js"
```

---

### Task 10: Wire everything together in App.tsx

**Files:**
- Modify: `src/taskpane/App.tsx`

- [ ] **Step 1: Update App.tsx with full integration**

```tsx
import React, { useState } from "react";
import { ChatPanel } from "./components/ChatPanel";
import { SettingsPanel } from "./components/SettingsPanel";
import { sendChatRequest } from "./services/aiClient";
import { readDocumentStructure, applyAIResponse } from "./services/wordDocument";

type Tab = "chat" | "settings";

const SYSTEM_PROMPT = `You are a Word document writing assistant. The user will provide:
1. The current document structure (text and tables)
2. Raw data to fill into the document

Your job is to fill the document template with the provided data.

Respond with a JSON code block containing an array of write instructions:
\`\`\`json
[
  {"type": "table_cell", "tableIndex": 0, "row": 1, "col": 1, "value": "content"},
  {"type": "replace", "searchText": "placeholder text", "value": "replacement"},
  {"type": "paragraph", "value": "new paragraph text"}
]
\`\`\`

Rules:
- tableIndex, row, col are 0-based
- Use "replace" to find and replace specific text in the document
- Use "table_cell" to fill specific cells in tables
- Use "paragraph" to add new paragraphs
- Only output the JSON block, no other text`;

export function App() {
  const [activeTab, setActiveTab] = useState<Tab>("chat");

  const handleSend = async (userMessage: string): Promise<string> => {
    // Read current document structure
    let docStructure = "";
    try {
      docStructure = await readDocumentStructure();
    } catch {
      docStructure = "(Could not read document — running outside Word)";
    }

    const fullMessage = `## Current Document Structure:\n${docStructure}\n\n## Raw Data from User:\n${userMessage}`;

    // Send to AI
    const aiResponse = await sendChatRequest(SYSTEM_PROMPT, fullMessage);

    // Try to apply response to document
    try {
      await applyAIResponse(aiResponse);
      return "Document updated successfully.\n\nAI Response:\n" + aiResponse;
    } catch {
      return "AI Response (could not auto-apply):\n" + aiResponse;
    }
  };

  return (
    <div className="app">
      <header className="app-header">
        <button
          className={`tab-btn ${activeTab === "chat" ? "active" : ""}`}
          onClick={() => setActiveTab("chat")}
        >
          Chat
        </button>
        <button
          className={`tab-btn ${activeTab === "settings" ? "active" : ""}`}
          onClick={() => setActiveTab("settings")}
        >
          Settings
        </button>
      </header>
      <main className="app-content">
        {activeTab === "chat" ? (
          <ChatPanel onSend={handleSend} />
        ) : (
          <SettingsPanel />
        )}
      </main>
    </div>
  );
}
```

- [ ] **Step 2: Verify in browser** (AI call without Word context)

Run dev server, configure settings with a valid API URL, send a message. Should get AI response (document apply will gracefully fail in browser).

- [ ] **Step 3: Verify in Mac Word** (full integration)

Sideload manifest.xml in Mac Word with a template document open. Send raw data. AI should respond and modify the document.

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/App.tsx
git commit -m "feat: integrate AI client and Word document service in main app"
```

---

## Chunk 6: Polish & Package for Deployment

### Task 11: Add npm scripts and production build

**Files:**
- Modify: `package.json`

- [ ] **Step 1: Add scripts to package.json**

```json
{
  "scripts": {
    "dev": "webpack serve --mode development",
    "build": "webpack --mode production",
    "start": "webpack serve --mode development"
  }
}
```

- [ ] **Step 2: Verify dev script works**

Run: `npm run dev`
Expected: Dev server starts at https://localhost:3000

- [ ] **Step 3: Verify production build**

Run: `npm run build`
Expected: `dist/` folder created with `taskpane.html` and `taskpane.bundle.js`

- [ ] **Step 4: Commit**

```bash
git add package.json
git commit -m "chore: add dev and build npm scripts"
```

---

### Task 12: End-to-end test with Mac Word

- [ ] **Step 1: Start dev server**

```bash
npm run dev
```

- [ ] **Step 2: Create a test template document**

Open Word, create a document with:
- Title: "Q4 2025 Financial Report"
- A table with headers: Metric | Q3 | Q4 | Change
- Empty rows for: Revenue, Operating Profit, Net Income

- [ ] **Step 3: Sideload add-in**

Insert > Add-ins > Upload manifest.xml

- [ ] **Step 4: Configure AI server**

In Settings tab, enter AI server URL and model name.

- [ ] **Step 5: Test document fill**

In Chat tab, paste:
```
Revenue: Q3 $2.5B, Q4 $3.1B
Operating Profit: Q3 $500M, Q4 $700M
Net Income: Q3 $350M, Q4 $520M
```

Verify: AI reads document structure, generates fill instructions, and updates the table cells.
