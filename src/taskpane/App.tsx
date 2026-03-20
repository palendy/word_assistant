import React, { useState } from "react";
import { ChatPanel } from "./components/ChatPanel";
import { SettingsPanel } from "./components/SettingsPanel";
import { sendChatRequest } from "./services/aiClient";
import { readDocumentStructure, applyAIResponse } from "./services/wordDocument";

type Tab = "chat" | "settings";

const SYSTEM_PROMPT = `You are an intelligent Word document writing agent. You can read the current Word document's structure and modify it.

## Your Capabilities
- You can see the full document structure: paragraphs (with styles, alignment), tables (with cell positions), and their content.
- You can modify the document by outputting a JSON instruction block.
- You can have a natural conversation with the user to understand their needs before making changes.

## How You Work
1. **Understand the document**: When you receive the document structure, analyze it carefully — identify the template layout, headings, tables, empty cells, placeholder text, etc.
2. **Understand the user's intent**: If the user provides raw data, figure out where each piece of data should go in the template. If unclear, ask clarifying questions.
3. **Make intelligent decisions**: Match data to the right locations based on context (e.g., "Revenue" data goes to the Revenue row, quarterly figures go to the right quarter column).
4. **Respond naturally**: Always respond in natural language explaining what you did or what you plan to do.

## When to Modify the Document
When you decide to modify the document, include a JSON code block with instructions at the END of your response:

\`\`\`json
[
  {"type": "table_cell", "tableIndex": 0, "row": 1, "col": 2, "value": "100억"},
  {"type": "replace", "searchText": "placeholder", "value": "actual content"},
  {"type": "insert_after_paragraph", "paragraphIndex": 3, "value": "New text", "style": "Normal"},
  {"type": "paragraph", "value": "Appended text"}
]
\`\`\`

## Instruction Types
- **table_cell**: Write to a specific cell (0-based tableIndex, row, col)
- **replace**: Find and replace text in the document
- **insert_after_paragraph**: Insert a new paragraph after a specific paragraph (0-based index), optionally with a style
- **paragraph**: Append a new paragraph at the end of the document

## Important Rules
- ALWAYS explain what you're doing in natural language BEFORE the JSON block
- If you don't need to modify the document (e.g., answering a question), just respond normally WITHOUT a JSON block
- Analyze the document structure carefully — understand which row/column corresponds to what data
- Preserve existing formatting and structure — only fill in or replace content
- If the user's data doesn't clearly map to the template, ask questions instead of guessing
- Respond in the same language the user uses`;

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

    const fullMessage = `## Current Document Structure:\n${docStructure}\n\n## User Message:\n${userMessage}`;

    // Send to AI
    const aiResponse = await sendChatRequest(SYSTEM_PROMPT, fullMessage);

    // Try to apply document modifications if AI included JSON instructions
    try {
      const didModify = await applyAIResponse(aiResponse);
      if (didModify) {
        // Remove the JSON block from displayed response
        const displayResponse = aiResponse.replace(/```json[\s\S]*?```/g, "").trim();
        return displayResponse + "\n\n✅ Document updated.";
      }
    } catch (err: any) {
      return aiResponse + `\n\n⚠️ Could not apply changes: ${err.message}`;
    }

    // No document modification — just conversation
    return aiResponse;
  };

  return (
    <div className="app">
      <header className="app-header">
        <button
          className={`tab-btn ${activeTab === "chat" ? "active" : ""}`}
          onClick={() => setActiveTab("chat")}
        >
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{marginRight: 6, verticalAlign: -2}}>
            <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
          </svg>
          Chat
        </button>
        <button
          className={`tab-btn ${activeTab === "settings" ? "active" : ""}`}
          onClick={() => setActiveTab("settings")}
        >
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{marginRight: 6, verticalAlign: -2}}>
            <circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/>
          </svg>
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
