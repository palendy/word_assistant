import React, { useState, useRef, useCallback } from "react";
import { ChatPanel, AgentStep } from "./components/ChatPanel";
import { SettingsPanel } from "./components/SettingsPanel";
import { ChatMessage, sendChatStreamRequest } from "./services/aiClient";
import { TOOL_DEFINITIONS, executeTool, undoLastOperation, clearUndoStack } from "./services/tools";

type Tab = "chat" | "settings";

const MAX_AGENT_LOOPS = 10;

const SYSTEM_PROMPT = `You are a powerful Word document automation agent. You MUST use tools to modify the document — NEVER output markdown tables, code blocks, or raw text as a substitute.

## Tools (8 total)
| Tool | Use for |
|------|---------|
| read_document | Read document structure (ALWAYS call first) |
| batch_write | Write paragraphs (with styles). Use for all text creation |
| replace_text | Find and replace text |
| insert_table | Create new tables |
| write_table_cells | Edit existing table cells |
| clear_document | Clear all content |
| insert_ooxml | Insert complex formatted content (OOXML) |
| execute_word_js | **Everything else** — formatting, resizing, borders, headers, footers, page breaks, deleting, fonts, alignment, margins, column widths, etc. |

## Rules
1. ALWAYS call \`read_document\` first if you haven't seen the document yet.
2. NEVER output markdown tables — use \`insert_table\`.
3. NEVER modify cell content to change table width — use \`execute_word_js\` to set table.width or autoFitWindow().
4. For text creation, prefer \`batch_write\` (handles multiple paragraphs in one call).
5. For ANY formatting/layout/structural operation, use \`execute_word_js\`. It runs inside Word.run().
6. Respond in the same language the user uses. Be concise.

## execute_word_js Examples
Table width: \`const t = context.document.body.tables; t.load("items"); await context.sync(); t.items[0].width = 300;\`
Delete paragraph: \`const p = context.document.body.paragraphs; p.load("items"); await context.sync(); p.items[5].delete();\`
Bold text: \`const r = context.document.body.search("Title"); r.load("items"); await context.sync(); r.items[0].font.bold = true;\`
Header: \`context.document.sections.getFirst().getHeader("Primary").insertParagraph("My Header", "Start");\`
Page break: \`context.document.body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);\`
Autofit table: \`const t = context.document.body.tables; t.load("items"); await context.sync(); t.items[0].autoFitWindow();\`

ALWAYS load("items") and await context.sync() before accessing .items[].`;

export function App() {
  const [activeTab, setActiveTab] = useState<Tab>("chat");
  const conversationHistory = useRef<ChatMessage[]>([]);

  const handleClear = useCallback(() => {
    conversationHistory.current = [];
    clearUndoStack();
  }, []);

  const handleUndo = useCallback(async (): Promise<string> => {
    try {
      return await undoLastOperation();
    } catch (err: any) {
      return `Undo failed: ${err.message}`;
    }
  }, []);

  const handleSend = useCallback(
    (
      userMessage: string,
      onToken: (token: string) => void,
      onDone: (fullResponse: string, thinking?: string, steps?: AgentStep[]) => void,
      onError: (error: Error) => void,
      onStepUpdate: (steps: AgentStep[]) => void
    ) => {
      (async () => {
        conversationHistory.current.push({
          role: "user",
          content: userMessage,
        });

        const messages: ChatMessage[] = [
          { role: "system", content: SYSTEM_PROMPT },
          ...conversationHistory.current,
        ];

        const allSteps: AgentStep[] = [];
        let finalContent = "";
        let loopCount = 0;

        try {
          // Agent loop: keep calling AI until no more tool calls
          while (loopCount < MAX_AGENT_LOOPS) {
            loopCount++;

            const result = await sendChatStreamRequest(
              messages,
              onToken,
              TOOL_DEFINITIONS
            );

            if (result.content) {
              finalContent += result.content;
            }

            // No tool calls — we're done
            if (!result.toolCalls || result.toolCalls.length === 0) {
              break;
            }

            // Add assistant message with tool calls to history
            const assistantMsg: ChatMessage = {
              role: "assistant",
              content: result.content || null,
              tool_calls: result.toolCalls,
            };
            messages.push(assistantMsg);
            conversationHistory.current.push(assistantMsg);

            // Execute each tool call
            for (const tc of result.toolCalls) {
              const step: AgentStep = {
                toolName: tc.function.name,
                toolCallId: tc.id,
                args: {},
                result: "",
                status: "running",
              };

              try {
                step.args = JSON.parse(tc.function.arguments);
              } catch {
                step.args = { raw: tc.function.arguments };
              }

              allSteps.push(step);
              onStepUpdate([...allSteps]);

              try {
                const toolResult = await executeTool(tc.function.name, step.args);
                step.result = toolResult.result;
                step.status = "done";
              } catch (err: any) {
                step.result = `Error: ${err.message}`;
                step.status = "error";
              }

              onStepUpdate([...allSteps]);

              // Add tool result to messages
              const toolMsg: ChatMessage = {
                role: "tool",
                content: step.result,
                tool_call_id: tc.id,
              };
              messages.push(toolMsg);
              conversationHistory.current.push(toolMsg);
            }

            // Clear streaming content for next round
            onToken("\n");
          }

          // Add final assistant content to history
          if (finalContent) {
            conversationHistory.current.push({
              role: "assistant",
              content: finalContent,
            });
          }

          // Clean up any thinking tags if model outputs them (optional)
          const displayContent = finalContent
            .replace(/<thinking>[\s\S]*?<\/thinking>/g, "")
            .trim();
          // No longer rely on <thinking> tags for UI — steps are the visible process
          const thinking = undefined;

          onDone(displayContent, thinking, allSteps.length > 0 ? allSteps : undefined);
        } catch (err: any) {
          onError(err);
        }
      })();
    },
    []
  );

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
          <ChatPanel onSend={handleSend} onClear={handleClear} onUndo={handleUndo} />
        ) : (
          <SettingsPanel />
        )}
      </main>
    </div>
  );
}
