import OpenAI from "openai";
import type { ChatCompletionMessageParam, ChatCompletionTool } from "openai/resources";
import { AISettings } from "../storage/settings";

export type Message = {
  role: "user" | "assistant" | "system" | "tool";
  content: string;
  tool_call_id?: string;
  name?: string;
};

export type ToolDefinition = {
  name: string;
  description: string;
  inputSchema: Record<string, unknown>;
};

export type ToolCall = {
  id: string;
  name: string;
  arguments: Record<string, unknown>;
};

export type StreamChunk =
  | { type: "text"; delta: string }
  | { type: "tool_call"; toolCall: ToolCall }
  | { type: "done" };

// Single unified streaming function using OpenAI-compatible API (works for
// OpenRouter, OpenAI, Gemini via OpenAI-compat endpoint, and custom providers).
export async function* streamChat(
  settings: AISettings,
  messages: Message[],
  tools: ToolDefinition[] = []
): AsyncGenerator<StreamChunk> {
  const client = new OpenAI({
    apiKey: settings.apiKey,
    baseURL: settings.baseUrl,
    dangerouslyAllowBrowser: true,
    defaultHeaders:
      settings.provider === "openrouter"
        ? {
            "HTTP-Referer": "https://github.com/askbard/mcpwordextension",
            "X-Title": "MCP Office Extension",
          }
        : {},
  });

  const oaiMessages = messages.map(
    (m): ChatCompletionMessageParam =>
      m.role === "tool"
        ? { role: "tool", content: m.content, tool_call_id: m.tool_call_id! }
        : m.role === "system"
        ? { role: "system", content: m.content }
        : m.role === "assistant"
        ? { role: "assistant", content: m.content }
        : { role: "user", content: m.content }
  );

  const oaiTools: ChatCompletionTool[] = tools.map((t) => ({
    type: "function",
    function: {
      name: t.name,
      description: t.description,
      parameters: t.inputSchema as Record<string, unknown>,
    },
  }));

  const stream = await client.chat.completions.create({
    model: settings.model,
    messages: oaiMessages,
    stream: true,
    ...(oaiTools.length > 0 ? { tools: oaiTools } : {}),
  });

  const pendingToolCalls: Map<number, { id: string; name: string; argsRaw: string }> = new Map();

  for await (const chunk of stream) {
    const choice = chunk.choices[0];
    if (!choice) continue;

    const delta = choice.delta;

    if (delta.content) {
      yield { type: "text", delta: delta.content };
    }

    if (delta.tool_calls) {
      for (const tc of delta.tool_calls) {
        if (!pendingToolCalls.has(tc.index)) {
          pendingToolCalls.set(tc.index, {
            id: tc.id ?? "",
            name: tc.function?.name ?? "",
            argsRaw: "",
          });
        }
        const existing = pendingToolCalls.get(tc.index)!;
        if (tc.id) existing.id = tc.id;
        if (tc.function?.name) existing.name += tc.function.name;
        if (tc.function?.arguments) existing.argsRaw += tc.function.arguments;
      }
    }

    if (choice.finish_reason === "tool_calls") {
      for (const [, tc] of pendingToolCalls) {
        let args: Record<string, unknown> = {};
        try {
          args = JSON.parse(tc.argsRaw);
        } catch {
          args = { _raw: tc.argsRaw };
        }
        yield { type: "tool_call", toolCall: { id: tc.id, name: tc.name, arguments: args } };
      }
      pendingToolCalls.clear();
    }

    if (choice.finish_reason === "stop") {
      yield { type: "done" };
    }
  }

  yield { type: "done" };
}
