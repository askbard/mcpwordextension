import { useRef, useCallback } from "react";
import { streamChat, Message, ToolDefinition, ToolCall } from "../../providers/index";
import { McpConnection, callMcpTool } from "../../mcp/client";
import { loadAISettings } from "../../storage/settings";

export type AgentMessage = {
  id: string;
  role: "user" | "assistant" | "tool_result";
  content: string;
  toolName?: string;
  isStreaming?: boolean;
};

type Opts = {
  connections: McpConnection[];
  onMessage: (msg: AgentMessage) => void;
  onUpdateLast: (id: string, delta: string, done?: boolean) => void;
};

let msgCounter = 0;
function nextId() {
  return `msg_${++msgCounter}_${Date.now()}`;
}

export function useAgentLoop({ connections, onMessage, onUpdateLast }: Opts) {
  const abortRef = useRef<boolean>(false);

  const run = useCallback(
    async (history: Message[]) => {
      abortRef.current = false;
      const settings = loadAISettings();

      if (!settings.apiKey) {
        onMessage({
          id: nextId(),
          role: "assistant",
          content: "⚠️ No API key set. Please configure one in the Settings tab.",
        });
        return;
      }

      const tools: ToolDefinition[] = connections.flatMap((c) => c.tools);

      // Add system prompt
      const messages: Message[] = [
        { role: "system", content: settings.systemPrompt },
        ...history,
      ];

      // Agentic loop — keep going until no more tool calls
      while (!abortRef.current) {
        const assistantId = nextId();
        onMessage({ id: assistantId, role: "assistant", content: "", isStreaming: true });

        let fullText = "";
        const pendingToolCalls: ToolCall[] = [];

        for await (const chunk of streamChat(settings, messages, tools)) {
          if (abortRef.current) break;

          if (chunk.type === "text") {
            fullText += chunk.delta;
            onUpdateLast(assistantId, chunk.delta);
          } else if (chunk.type === "tool_call") {
            pendingToolCalls.push(chunk.toolCall);
          } else if (chunk.type === "done") {
            break;
          }
        }

        onUpdateLast(assistantId, "", true);

        // Add assistant message to history
        messages.push({ role: "assistant", content: fullText });

        if (pendingToolCalls.length === 0 || abortRef.current) {
          break;
        }

        // Execute each tool call
        for (const tc of pendingToolCalls) {
          const toolResultId = nextId();
          onMessage({
            id: toolResultId,
            role: "tool_result",
            content: `Calling **${tc.name}**...`,
            toolName: tc.name,
          });

          let result = "";
          const conn = connections.find((c) =>
            c.tools.some((t) => t.name === tc.name)
          );

          if (conn) {
            try {
              result = await callMcpTool(conn, tc.name, tc.arguments);
            } catch (e) {
              result = `Error: ${e instanceof Error ? e.message : String(e)}`;
            }
          } else {
            result = `No MCP server found for tool: ${tc.name}`;
          }

          onUpdateLast(toolResultId, result, true);

          messages.push({
            role: "tool",
            content: result,
            tool_call_id: tc.id,
            name: tc.name,
          });
        }
      }
    },
    [connections, onMessage, onUpdateLast]
  );

  const abort = useCallback(() => {
    abortRef.current = true;
  }, []);

  return { run, abort };
}
