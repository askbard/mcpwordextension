import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import type { McpServer } from "../storage/settings";
import type { ToolDefinition } from "../providers/index";

export type McpConnection = {
  server: McpServer;
  client: Client;
  tools: ToolDefinition[];
};

export async function connectMcpServer(server: McpServer): Promise<McpConnection> {
  const transport = new StreamableHTTPClientTransport(new URL(server.url), {
    requestInit: {
      headers: server.headers,
    },
  });

  const client = new Client(
    { name: "mcp-office-extension", version: "1.0.0" }
  );

  await client.connect(transport);

  const listResult = await client.listTools();
  const tools = listResult.tools;

  const toolDefs: ToolDefinition[] = tools.map((t) => ({
    name: `${server.name}__${t.name}`,
    description: `[${server.name}] ${t.description ?? t.name}`,
    inputSchema: (t.inputSchema as Record<string, unknown>) ?? { type: "object", properties: {} },
  }));

  return { server, client, tools: toolDefs };
}

export async function callMcpTool(
  connection: McpConnection,
  toolName: string,
  args: Record<string, unknown>
): Promise<string> {
  // Strip the "serverName__" prefix to get the actual tool name
  const actualName = toolName.replace(`${connection.server.name}__`, "");

  const result = await connection.client.callTool({
    name: actualName,
    arguments: args,
  });

  const content = result.content as Array<{ type: string; text?: string; resource?: unknown }>;
  if (!content || content.length === 0) {
    return "(no result)";
  }

  return content
    .map((c) => {
      if (c.type === "text") return c.text ?? "";
      if (c.type === "resource") return JSON.stringify(c.resource);
      return JSON.stringify(c);
    })
    .join("\n");
}

export async function disconnectMcpServer(connection: McpConnection): Promise<void> {
  await connection.client.close();
}
