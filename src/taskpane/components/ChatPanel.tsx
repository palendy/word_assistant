import React, { useState, useRef, useEffect, useCallback } from "react";
import { Message, MessageBubble } from "./MessageBubble";

interface Command {
  name: string;
  description: string;
  action: () => string | null;
}

interface ChatPanelProps {
  onSend: (
    userMessage: string,
    onToken: (token: string) => void,
    onDone: (fullResponse: string) => void,
    onError: (error: Error) => void
  ) => void;
  onClear: () => void;
}

export function ChatPanel({ onSend, onClear }: ChatPanelProps) {
  const [messages, setMessages] = useState<Message[]>([]);
  const [streamingContent, setStreamingContent] = useState("");
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [showCommands, setShowCommands] = useState(false);
  const [commandFilter, setCommandFilter] = useState("");
  const [selectedCommandIndex, setSelectedCommandIndex] = useState(0);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const commands: Command[] = [
    {
      name: "/clear",
      description: "Clear all chat history",
      action: () => {
        setMessages([]);
        setStreamingContent("");
        onClear();
        return null;
      },
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
  }, [messages, streamingContent, loading]);

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

      onSend(
        text,
        // onToken
        (token) => {
          setStreamingContent((prev) => prev + token);
        },
        // onDone
        (fullResponse) => {
          setStreamingContent("");
          setMessages((prev) => [
            ...prev,
            { role: "assistant", content: fullResponse },
          ]);
          setLoading(false);
        },
        // onError
        (error) => {
          setStreamingContent("");
          setMessages((prev) => [
            ...prev,
            { role: "assistant", content: `Error: ${error.message}` },
          ]);
          setLoading(false);
        }
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

    const commandMessages: Record<string, string> = {
      "__CMD_SUMMARY__":
        "Read the current document and provide a concise summary of its content.",
      "__CMD_HELP__": "__HELP__",
    };

    const aiMessage = commandMessages[result];

    if (aiMessage === "__HELP__") {
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

    setMessages((prev) => [...prev, { role: "user", content: cmd.name }]);
    sendMessage(aiMessage);
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
          {streamingContent && (
            <div className="message assistant">
              <div className="message-role">AI Assistant</div>
              <div className="message-content">
                {streamingContent}
                <span className="streaming-cursor" />
              </div>
            </div>
          )}
          {loading && !streamingContent && (
            <div className="loading-indicator">
              <div className="dot" />
              <div className="dot" />
              <div className="dot" />
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
