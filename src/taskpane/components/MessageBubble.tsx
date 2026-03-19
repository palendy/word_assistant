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
