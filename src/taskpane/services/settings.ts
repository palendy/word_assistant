export interface Settings {
  apiUrl: string;
  modelName: string;
}

const STORAGE_KEY = "word-ai-assistant-settings";

const DEFAULT_SETTINGS: Settings = {
  apiUrl: "http://localhost:11434/v1/chat/completions",
  modelName: "llama3",
};

export function loadSettings(): Settings {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (stored) {
      return { ...DEFAULT_SETTINGS, ...JSON.parse(stored) };
    }
  } catch {
    // ignore parse errors
  }
  return { ...DEFAULT_SETTINGS };
}

export function saveSettings(settings: Settings): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}
