import { loadSettings } from "./settings";

interface ChatMessage {
  role: "system" | "user" | "assistant";
  content: string;
}

interface ChatCompletionResponse {
  choices: Array<{
    message: {
      role: string;
      content: string;
    };
  }>;
}

export async function sendChatRequest(
  systemPrompt: string,
  userMessage: string
): Promise<string> {
  const settings = loadSettings();
  const apiUrl = settings.apiUrl.trim();
  const apiKey = settings.apiKey.trim();
  const modelName = settings.modelName.trim();

  if (!apiUrl) {
    throw new Error("AI Server URL is not configured. Go to Settings tab.");
  }

  const messages: ChatMessage[] = [
    { role: "system", content: systemPrompt },
    { role: "user", content: userMessage },
  ];

  const headers: Record<string, string> = {
    "Content-Type": "application/json",
  };
  if (apiKey) {
    headers["Authorization"] = `Bearer ${apiKey}`;
  }

  const response = await fetch(apiUrl, {
    method: "POST",
    headers,
    body: JSON.stringify({
      model: modelName,
      messages,
      temperature: 0.3,
    }),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`API error (${response.status}): ${errorText}`);
  }

  const data: ChatCompletionResponse = await response.json();

  if (!data.choices || data.choices.length === 0) {
    throw new Error("No response from AI model");
  }

  return data.choices[0].message.content;
}
