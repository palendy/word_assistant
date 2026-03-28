import React, { useState, useRef, useEffect, useCallback } from "react";
import { Message, MessageBubble } from "./MessageBubble";
import { PlanCard } from "./PlanCard";
import { ProposedPlan } from "../App";

export interface AgentStep {
  toolName: string;
  toolCallId: string;
  args: Record<string, any>;
  result: string;
  status: "running" | "done" | "error";
}

interface Command {
  name: string;
  description: string;
  action: () => string | null;
}

interface ChatPanelProps {
  onSend: (
    userMessage: string,
    onToken: (token: string) => void,
    onDone: (fullResponse: string, thinking?: string, steps?: AgentStep[]) => void,
    onError: (error: Error) => void,
    onStepUpdate: (steps: AgentStep[]) => void,
    onPlanProposed: (plan: ProposedPlan) => Promise<boolean>
  ) => void;
  onClear: () => void;
  onUndo: () => Promise<string>;
}

const TOOL_LABELS: Record<string, string> = {
  read_document: "📄 Reading document",
  batch_write: "📝 Writing content",
  replace_text: "🔄 Replacing text",
  insert_table: "📊 Creating table",
  write_table_cells: "✏️ Editing table cells",
  clear_document: "🧹 Clearing document",
  insert_ooxml: "📄 Inserting formatted content",
  execute_word_js: "🔧 Executing Word API",
  propose_plan: "📋 Proposing plan",
};

function StepCard({ step }: { step: AgentStep }) {
  const label = TOOL_LABELS[step.toolName] || step.toolName;
  const icon =
    step.status === "running" ? "⏳" : step.status === "done" ? "✓" : "✗";
  const statusClass = `step-status-${step.status}`;

  let detail = "";
  if (step.toolName === "write_table_cells" && step.args.changes) {
    detail = `(${step.args.changes.length} cells)`;
  } else if (step.toolName === "replace_text") {
    detail = `"${step.args.searchText}"`;
  } else if (step.toolName === "execute_word_js" && step.args.description) {
    detail = `— ${step.args.description}`;
  } else if (step.toolName === "batch_write" && step.args.paragraphs) {
    detail = `(${step.args.paragraphs.length} paragraphs)`;
  } else if (step.toolName === "insert_ooxml" && step.args.description) {
    detail = `— ${step.args.description}`;
  }

  return (
    <div className={`step-card ${statusClass}`}>
      <span className="step-icon">{icon}</span>
      <span className="step-label">
        {label} {detail}
      </span>
    </div>
  );
}

export function ChatPanel({ onSend, onClear, onUndo }: ChatPanelProps) {
  const [messages, setMessages] = useState<Message[]>([]);
  const [streamingContent, setStreamingContent] = useState("");
  const [currentSteps, setCurrentSteps] = useState<AgentStep[]>([]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [showCommands, setShowCommands] = useState(false);
  const [commandFilter, setCommandFilter] = useState("");
  const [selectedCommandIndex, setSelectedCommandIndex] = useState(0);
  const [pendingPlan, setPendingPlan] = useState<ProposedPlan | null>(null);
  const planResolverRef = useRef<((approved: boolean) => void) | null>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const handlePlanProposed = useCallback((plan: ProposedPlan): Promise<boolean> => {
    return new Promise((resolve) => {
      planResolverRef.current = resolve;
      setPendingPlan(plan);
    });
  }, []);

  const handlePlanApprove = useCallback(() => {
    setPendingPlan(null);
    planResolverRef.current?.(true);
    planResolverRef.current = null;
  }, []);

  const handlePlanCancel = useCallback(() => {
    setPendingPlan(null);
    planResolverRef.current?.(false);
    planResolverRef.current = null;
  }, []);

  const commands: Command[] = [
    {
      name: "/clear",
      description: "Clear all chat history",
      action: () => {
        setMessages([]);
        setStreamingContent("");
        setCurrentSteps([]);
        onClear();
        return null;
      },
    },
    {
      name: "/undo",
      description: "Undo last document modification",
      action: () => "__CMD_UNDO__",
    },
    {
      name: "/summary",
      description: "Summarize the current document",
      action: () => "__CMD_SUMMARY__",
    },
    {
      name: "/help",
      description: "Show available commands",
      action: () => "__CMD_HELP__",
    },
  ];

  const filteredCommands = commands.filter((cmd) =>
    cmd.name.toLowerCase().includes(commandFilter.toLowerCase())
  );

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, streamingContent, loading, currentSteps]);

  useEffect(() => {
    setSelectedCommandIndex(0);
  }, [commandFilter]);

  const autoResize = () => {
    const el = textareaRef.current;
    if (el) {
      el.style.height = "auto";
      el.style.height = Math.min(el.scrollHeight, 120) + "px";
    }
  };

  const handleInputChange = (value: string) => {
    setInput(value);
    if (value.startsWith("/")) {
      setShowCommands(true);
      setCommandFilter(value);
    } else {
      setShowCommands(false);
      setCommandFilter("");
    }
    autoResize();
  };

  const sendMessage = useCallback(
    (text: string) => {
      setLoading(true);
      setStreamingContent("");
      setCurrentSteps([]);

      onSend(
        text,
        (token) => {
          setStreamingContent((prev) => prev + token);
        },
        (fullResponse, thinking, steps) => {
          setStreamingContent("");
          setCurrentSteps([]);
          setPendingPlan(null);
          setMessages((prev) => [
            ...prev,
            {
              role: "assistant",
              content: fullResponse,
              thinking,
              steps,
            },
          ]);
          setLoading(false);
        },
        (error) => {
          setStreamingContent("");
          setCurrentSteps([]);
          setPendingPlan(null);
          setMessages((prev) => [
            ...prev,
            { role: "assistant", content: `Error: ${error.message}` },
          ]);
          setLoading(false);
        },
        (steps) => {
          setCurrentSteps([...steps]);
        },
        handlePlanProposed
      );
    },
    [onSend]
  );

  const executeCommand = (cmd: Command) => {
    setShowCommands(false);
    setInput("");
    setCommandFilter("");

    const result = cmd.action();
    if (result === null) return;

    if (result === "__CMD_UNDO__") {
      setMessages((prev) => [...prev, { role: "user", content: "/undo" }]);
      onUndo().then((msg) => {
        setMessages((prev) => [...prev, { role: "assistant", content: msg }]);
      });
      return;
    }

    if (result === "__CMD_HELP__") {
      const helpText = commands
        .map((c) => `${c.name}  —  ${c.description}`)
        .join("\n");
      setMessages((prev) => [
        ...prev,
        {
          role: "assistant",
          content: `Available commands:\n\n${helpText}\n\nType / to see the command list.`,
        },
      ]);
      return;
    }

    const commandMessages: Record<string, string> = {
      "__CMD_SUMMARY__":
        "Read the current document and provide a concise summary of its content.",
    };

    const aiMessage = commandMessages[result];
    if (aiMessage) {
      setMessages((prev) => [...prev, { role: "user", content: cmd.name }]);
      sendMessage(aiMessage);
    }
  };

  const handleSend = async () => {
    const text = input.trim();
    if (!text || loading) return;

    if (text.startsWith("/")) {
      const cmd = commands.find(
        (c) => c.name.toLowerCase() === text.toLowerCase()
      );
      if (cmd) {
        executeCommand(cmd);
        return;
      }
    }

    setInput("");
    setShowCommands(false);
    if (textareaRef.current) textareaRef.current.style.height = "auto";
    setMessages((prev) => [...prev, { role: "user", content: text }]);
    sendMessage(text);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (showCommands && filteredCommands.length > 0) {
      if (e.key === "ArrowDown") {
        e.preventDefault();
        setSelectedCommandIndex((prev) =>
          prev < filteredCommands.length - 1 ? prev + 1 : 0
        );
        return;
      }
      if (e.key === "ArrowUp") {
        e.preventDefault();
        setSelectedCommandIndex((prev) =>
          prev > 0 ? prev - 1 : filteredCommands.length - 1
        );
        return;
      }
      if (e.key === "Tab" || (e.key === "Enter" && !e.shiftKey)) {
        e.preventDefault();
        executeCommand(filteredCommands[selectedCommandIndex]);
        return;
      }
      if (e.key === "Escape") {
        setShowCommands(false);
        return;
      }
    }

    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  // Clean streaming content of any thinking tags
  const displayStreamingContent = streamingContent
    .replace(/<thinking>[\s\S]*?<\/thinking>/g, "")
    .replace(/<thinking>[\s\S]*/g, "")
    .trim();

  // Show status based on actual agent state
  const isWaitingForAI = loading && !streamingContent && currentSteps.length === 0 && !pendingPlan;
  const isProcessingNextStep = loading && !streamingContent && currentSteps.length > 0 && !currentSteps.some(s => s.status === "running") && !pendingPlan;

  return (
    <div className="chat-panel">
      {messages.length === 0 && !loading && !streamingContent ? (
        <div className="welcome">
          <div className="welcome-icon">A</div>
          <div className="welcome-title">AI Assistant</div>
          <div className="welcome-subtitle">
            Paste your raw data and I'll fill in the document template for you.
          </div>
          <div className="welcome-hint">
            Type <span className="hint-slash">/</span> for commands
          </div>
        </div>
      ) : (
        <div className="messages-list">
          {messages.map((msg, i) => (
            <MessageBubble key={i} {...msg} />
          ))}

          {/* Active step cards */}
          {currentSteps.length > 0 && (
            <div className="steps-container">
              {currentSteps.map((step, i) => (
                <StepCard key={i} step={step} />
              ))}
            </div>
          )}

          {/* Plan approval card — pauses agent loop */}
          {pendingPlan && (
            <PlanCard
              plan={pendingPlan}
              onApprove={handlePlanApprove}
              onCancel={handlePlanCancel}
            />
          )}

          {/* Streaming text content */}
          {displayStreamingContent && (
            <div className="message assistant">
              <div className="message-role">AI Assistant</div>
              <div className="message-content">
                {displayStreamingContent}
                <span className="streaming-cursor" />
              </div>
            </div>
          )}

          {/* Status indicators based on agent state */}
          {isWaitingForAI && (
            <div className="thinking-indicator">
              <span className="thinking-spinner" />
              <span>요청을 분석하고 있습니다...</span>
            </div>
          )}
          {isProcessingNextStep && (
            <div className="thinking-indicator">
              <span className="thinking-spinner" />
              <span>다음 단계를 계획하고 있습니다...</span>
            </div>
          )}
          <div ref={messagesEndRef} />
        </div>
      )}
      <div className="chat-input-area">
        {showCommands && filteredCommands.length > 0 && (
          <div className="command-palette">
            {filteredCommands.map((cmd, i) => (
              <div
                key={cmd.name}
                className={`command-item ${i === selectedCommandIndex ? "selected" : ""}`}
                onClick={() => executeCommand(cmd)}
                onMouseEnter={() => setSelectedCommandIndex(i)}
              >
                <span className="command-name">{cmd.name}</span>
                <span className="command-desc">{cmd.description}</span>
              </div>
            ))}
          </div>
        )}
        <div className="chat-input-wrapper">
          <textarea
            ref={textareaRef}
            className="chat-input"
            value={input}
            onChange={(e) => handleInputChange(e.target.value)}
            onKeyDown={handleKeyDown}
            placeholder="Type / for commands or paste your data..."
            rows={1}
          />
          <button
            className="send-btn"
            onClick={handleSend}
            disabled={loading || !input.trim()}
          >
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <line x1="22" y1="2" x2="11" y2="13" />
              <polygon points="22 2 15 22 11 13 2 9 22 2" />
            </svg>
          </button>
        </div>
      </div>
    </div>
  );
}
