import React, { useState } from "react";
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
      <div className="settings-title">Configuration</div>
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
        <label className="setting-label">API Key</label>
        <input
          className="setting-input"
          type="password"
          value={settings.apiKey}
          onChange={(e) => setSettings({ ...settings, apiKey: e.target.value })}
          placeholder="Optional — leave empty if not required"
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
      <button
        className={`save-btn ${saved ? "saved" : ""}`}
        onClick={handleSave}
      >
        {saved ? "Saved!" : "Save Settings"}
      </button>
    </div>
  );
}
