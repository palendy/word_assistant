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
