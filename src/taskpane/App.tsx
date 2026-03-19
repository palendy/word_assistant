import React, { useState } from "react";
import { ChatPanel } from "./components/ChatPanel";
import { SettingsPanel } from "./components/SettingsPanel";

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
          <ChatPanel onSend={async (msg) => `Echo: ${msg}`} />
        ) : (
          <SettingsPanel />
        )}
      </main>
    </div>
  );
}
