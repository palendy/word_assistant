import { loadSettings } from "./settings";

export interface ChatMessage {
  role: "system" | "user" | "assistant" | "tool";
  content: string | null;
  tool_calls?: ToolCallResponse[];
  tool_call_id?: string;
}

export interface ToolCallResponse {
  id: string;
  type: "function";
  function: {
    name: string;
    arguments: string;
  };
}

export interface StreamResult {
  content: string;
  toolCalls: ToolCallResponse[] | null;
}

const MAX_RETRIES = 3;
const BASE_RETRY_DELAY = 2000; // 2 seconds

export async function sendChatStreamRequest(
  messages: ChatMessage[],
  onToken: (token: string) => void,
  tools?: any[]
): Promise<StreamResult> {
  const settings = loadSettings();
  const apiUrl = settings.apiUrl.trim();
  const apiKey = settings.apiKey.trim();
  const modelName = settings.modelName.trim();

  if (!apiUrl) {
    throw new Error("AI Server URL is not configured. Go to Settings tab.");
  }

  const headers: Record<string, string> = {
    "Content-Type": "application/json",
  };
  if (apiKey) {
    headers["Authorization"] = `Bearer ${apiKey}`;
  }

  const body: any = {
    model: modelName,
    messages,
    temperature: 0.3,
    stream: true,
  };
  if (tools && tools.length > 0) {
    body.tools = tools;
  }

  // Retry loop for transient errors (429, 500, 502, 503, 504)
  let lastError: Error | null = null;
  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    try {
      const response = await fetch(apiUrl, {
        method: "POST",
        headers,
        body: JSON.stringify(body),
      });

      if (!response.ok) {
        const errorText = await response.text();
        const status = response.status;

        // Retryable errors
        if ((status === 429 || status >= 500) && attempt < MAX_RETRIES - 1) {
          // Parse retry-after from response if available
          let retryDelay = BASE_RETRY_DELAY * Math.pow(2, attempt);
          try {
            const errorJson = JSON.parse(errorText);
            const retryAfter =
              errorJson?.error?.metadata?.retry_after_seconds ||
              response.headers.get("retry-after");
            if (retryAfter) {
              retryDelay = Math.min(Number(retryAfter) * 1000, 60000);
            }
          } catch {
            // ignore parse error
          }

          onToken(`\n⏳ Rate limited (${status}). Retrying in ${Math.round(retryDelay / 1000)}s... (${attempt + 1}/${MAX_RETRIES})\n`);
          await sleep(retryDelay);
          continue;
        }

        throw new Error(`API error (${status}): ${errorText}`);
      }

      // Success — stream the response
      return await streamResponse(response, onToken);
    } catch (err: any) {
      lastError = err;

      // Network errors are retryable
      if (err.name === "TypeError" && attempt < MAX_RETRIES - 1) {
        const retryDelay = BASE_RETRY_DELAY * Math.pow(2, attempt);
        onToken(`\n⏳ Network error. Retrying in ${Math.round(retryDelay / 1000)}s...\n`);
        await sleep(retryDelay);
        continue;
      }

      throw err;
    }
  }

  throw lastError || new Error("Max retries exceeded");
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function streamResponse(
  response: Response,
  onToken: (token: string) => void
): Promise<StreamResult> {
  const reader = response.body?.getReader();
  if (!reader) throw new Error("No response stream");

  const decoder = new TextDecoder();
  let fullContent = "";
  let buffer = "";

  // Accumulate tool calls from stream deltas
  const toolCallMap = new Map<number, { id: string; name: string; arguments: string }>();

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;

    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split("\n");
    buffer = lines.pop() || "";

    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed || !trimmed.startsWith("data: ")) continue;

      const data = trimmed.slice(6);
      if (data === "[DONE]") break;

      try {
        const parsed = JSON.parse(data);
        const choice = parsed.choices?.[0];
        if (!choice) continue;

        const delta = choice.delta;
        if (!delta) continue;

        // Accumulate text content
        if (delta.content) {
          fullContent += delta.content;
          onToken(delta.content);
        }

        // Accumulate tool calls
        if (delta.tool_calls) {
          for (const tc of delta.tool_calls) {
            const idx = tc.index ?? 0;
            if (!toolCallMap.has(idx)) {
              toolCallMap.set(idx, { id: "", name: "", arguments: "" });
            }
            const acc = toolCallMap.get(idx)!;
            if (tc.id) acc.id = tc.id;
            if (tc.function?.name) acc.name = tc.function.name;
            if (tc.function?.arguments) acc.arguments += tc.function.arguments;
          }
        }
      } catch {
        // skip malformed JSON chunks
      }
    }
  }

  // Convert accumulated tool calls
  let toolCalls: ToolCallResponse[] | null = null;
  if (toolCallMap.size > 0) {
    toolCalls = [];
    for (const [, tc] of toolCallMap) {
      toolCalls.push({
        id: tc.id,
        type: "function",
        function: { name: tc.name, arguments: tc.arguments },
      });
    }
  }

  return { content: fullContent, toolCalls };
}
