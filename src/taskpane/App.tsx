import React, { useState, useRef, useCallback } from "react";
import { ChatPanel, AgentStep } from "./components/ChatPanel";
import { SettingsPanel } from "./components/SettingsPanel";
import { ChatMessage, sendChatStreamRequest } from "./services/aiClient";
import { TOOL_DEFINITIONS, executeTool, undoLastOperation, clearUndoStack } from "./services/tools";

export interface PlanStep {
  step_number: number;
  action: string;
  detail?: string;
}

export interface ProposedPlan {
  summary: string;
  steps: PlanStep[];
}

type Tab = "chat" | "settings";

const MAX_AGENT_LOOPS = 10;

const SYSTEM_PROMPT = `You are an expert Word document automation agent. Your job is to understand the user's intent and modify the Word document accordingly.

## Core Workflow
1. **Read first** — ALWAYS call \`read_document\` at the start of a session. This returns paragraph text (with formatting: bold, italic, size, color), table values, headers/footers, and styles.
2. **Act** — Use the appropriate tool to make changes.
3. **Re-read when needed** — After complex modifications, call \`read_document\` again to verify your changes or to get updated paragraph/table indices.

## Tools (8 total, use in priority order)

### High-level tools (prefer these):
- **read_document** — Read full document structure. Shows formatting, styles, alignment, tables, headers/footers.
- **batch_write** — Create paragraphs (with styles). Use for ALL text creation. Supports insert-after-index.
- **replace_text** — Find and replace text in document body. Supports replaceAll.
- **insert_table** — Create a new table from headers + rows. Auto-applies style and bold headers.
- **write_table_cells** — Edit specific cells in existing tables by (tableIndex, row, col, value).

### Low-level tools (use when high-level tools can't do it):
- **execute_word_js** — Run arbitrary Word JS API code. Use for: table resizing, borders, fonts, alignment, headers/footers, page breaks, deletion, cell shading, column widths, margins, etc.
- **insert_ooxml** — Insert Office Open XML for complex formatting (colored text, mixed inline formatting, images via base64).
- **clear_document** — Clear ALL content. Use with extreme caution.

## Rules
1. NEVER output markdown tables, code blocks, or formatted text as chat — ALWAYS use tools to write into the document.
2. NEVER modify cell text content to change table width — use \`execute_word_js\` with \`table.width\` or \`autoFitWindow()\`.
3. If the document already matches the user's request, say so — don't make unnecessary changes.
4. Respond in the same language the user uses. Be concise — let the document changes speak for themselves.
5. When user refers to visual elements (colors, bold, font sizes), use what \`read_document\` returns to understand the current state.

## execute_word_js Patterns (MUST follow)
Code runs inside \`Word.run(async (context) => { YOUR CODE })\`. You have \`context\` and \`Word\`.

**CRITICAL**: Always load collections before accessing items:
\`\`\`
const tables = context.document.body.tables;
tables.load("items");
await context.sync();
// NOW you can access tables.items[0]
\`\`\`

Common operations:
- Table width: \`tables.items[0].width = 300;\` or \`tables.items[0].autoFitWindow();\`
- Delete paragraph: \`paragraphs.items[5].delete();\`
- Bold text: \`const r = context.document.body.search("Title"); r.load("items"); await context.sync(); r.items[0].font.bold = true;\`
- Alignment: \`paragraphs.items[0].alignment = Word.Alignment.centered;\`
- Header: \`context.document.sections.getFirst().getHeader("Primary").insertParagraph("Header", "Start");\`
- Footer: \`context.document.sections.getFirst().getFooter("Primary").insertParagraph("Footer", "Start");\`
- Page break: \`context.document.body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);\`
- Cell background: \`table.getCell(0,0).shadingColor = "#4472C4";\`
- Column width: \`table.getColumn(0).preferredWidth = 100;\`

## When to Use propose_plan

Call **propose_plan BEFORE doing anything** when the request is complex:
- Multi-step restructuring (reorder sections, reformat the whole document)
- Writing long content (full reports, multiple sections, 5+ paragraphs)
- Cascading changes that touch many parts of the document
- Anything that could significantly alter the document structure

Call tools directly (NO propose_plan) for simple requests:
- Edit a cell → write_table_cells directly
- Replace text → replace_text directly
- Add one paragraph → batch_write directly
- Read the document → read_document directly
- Any clearly single-tool operation

When you call propose_plan:
1. List steps in execution order — what you will actually DO, not implementation details
2. Keep each action under 10 words (e.g. "Create title paragraph", "Insert 5-row table")
3. WAIT for approval — do not call any other tool in the same turn
4. After approval: execute step-by-step exactly as listed
5. After cancellation: acknowledge and stop`;

// Rough token estimate for conversation pruning
const MAX_HISTORY_CHARS = 100000; // ~25k tokens

function pruneHistory(history: ChatMessage[]): ChatMessage[] {
  let totalChars = 0;
  // Keep messages from the end (most recent first)
  const keep: ChatMessage[] = [];
  for (let i = history.length - 1; i >= 0; i--) {
    const msg = history[i];
    const msgChars = (msg.content || "").length + JSON.stringify(msg.tool_calls || "").length;
    totalChars += msgChars;
    if (totalChars > MAX_HISTORY_CHARS) {
      // Insert a summary message at the beginning
      keep.unshift({
        role: "system" as const,
        content: "[Earlier conversation history was pruned to save tokens. The user has been working with this document in prior turns.]",
      });
      break;
    }
    keep.unshift(msg);
  }
  return keep;
}

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
      onStepUpdate: (steps: AgentStep[]) => void,
      onPlanProposed: (plan: ProposedPlan) => Promise<boolean>
    ) => {
      (async () => {
        conversationHistory.current.push({
          role: "user",
          content: userMessage,
        });

        // Prune old history to prevent token overflow
        const prunedHistory = pruneHistory(conversationHistory.current);
        const messages: ChatMessage[] = [
          { role: "system", content: SYSTEM_PROMPT },
          ...prunedHistory,
        ];

        const allSteps: AgentStep[] = [];
        let finalContent = "";
        let loopCount = 0;
        let planWasCancelled = false;

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
              // Intercept propose_plan before executeTool
              if (tc.function.name === "propose_plan") {
                let planArgs: ProposedPlan;
                try {
                  planArgs = JSON.parse(tc.function.arguments);
                } catch {
                  planArgs = { summary: "계획을 파싱할 수 없습니다.", steps: [] };
                }

                const approved = await onPlanProposed(planArgs);

                const toolMsg: ChatMessage = {
                  role: "tool",
                  content: approved
                    ? "Plan approved by user. Proceed with execution step by step."
                    : "Plan cancelled by user. Do not proceed with any document changes.",
                  tool_call_id: tc.id,
                };
                messages.push(toolMsg);
                conversationHistory.current.push(toolMsg);

                if (!approved) {
                  planWasCancelled = true;
                  break;
                }
                continue;
              }

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

            if (planWasCancelled) break;

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
