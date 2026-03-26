import React, { useState } from "react";
import { AgentStep } from "./ChatPanel";

export interface Message {
  role: "user" | "assistant";
  content: string;
  thinking?: string;
  steps?: AgentStep[];
}

function ThinkingBlock({ content }: { content: string }) {
  const [expanded, setExpanded] = useState(false);

  return (
    <div className="thinking-block" onClick={() => setExpanded(!expanded)}>
      <div className="thinking-header">
        <span className="thinking-icon">💭</span>
        <span className="thinking-label">Thinking process</span>
        <span className="thinking-toggle">{expanded ? "▲" : "▼"}</span>
      </div>
      {expanded && (
        <div className="thinking-content">{content}</div>
      )}
    </div>
  );
}

function StepsSummary({ steps }: { steps: AgentStep[] }) {
  const [expanded, setExpanded] = useState(false);
  const writeSteps = steps.filter((s) => s.toolName !== "read_document");
  const totalChanges = writeSteps.length;

  if (totalChanges === 0) return null;

  return (
    <div className="steps-summary">
      <div
        className="steps-summary-header"
        onClick={() => setExpanded(!expanded)}
      >
        <span>✅ {totalChanges} modification(s) applied</span>
        <span className="thinking-toggle">{expanded ? "▲" : "▼"}</span>
      </div>
      {expanded && (
        <div className="steps-summary-list">
          {steps.map((step, i) => (
            <div key={i} className="steps-summary-item">
              <span className={`step-dot ${step.status}`} />
              <span>
                {step.toolName === "read_document"
                  ? "Read document"
                  : step.toolName === "write_table_cells"
                  ? `Wrote ${step.args.changes?.length || 0} cell(s)`
                  : step.toolName === "replace_text"
                  ? `Replaced "${step.args.searchText}"`
                  : step.toolName === "insert_paragraph"
                  ? "Inserted paragraph"
                  : step.toolName === "append_paragraph"
                  ? "Appended paragraph"
                  : step.toolName}
              </span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export function MessageBubble({ role, content, thinking, steps }: Message) {
  return (
    <div className={`message ${role}`}>
      <div className="message-role">
        {role === "user" ? "You" : "AI Assistant"}
      </div>
      {role === "assistant" && thinking && (
        <ThinkingBlock content={thinking} />
      )}
      {role === "assistant" && steps && steps.length > 0 && (
        <StepsSummary steps={steps} />
      )}
      <div className="message-content">{content}</div>
    </div>
  );
}
