export type AIProvider = "openrouter" | "openai" | "openai-compat" | "google";

export interface AISettings {
  provider: AIProvider;
  apiKey: string;
  baseUrl: string;
  model: string;
  systemPrompt: string;
}

export interface McpServer {
  id: string;
  name: string;
  url: string;
  headers: Record<string, string>;
  enabled: boolean;
}

const KEYS = {
  ai: "mcp_office_ai_settings",
  mcp: "mcp_office_mcp_servers",
} as const;

export const DEFAULT_AI_SETTINGS: AISettings = {
  provider: "openrouter",
  apiKey: "",
  baseUrl: "https://openrouter.ai/api/v1",
  model: "anthropic/claude-3.5-sonnet",
  systemPrompt: "You are a helpful AI assistant integrated into Microsoft Office.",
};

export function loadAISettings(): AISettings {
  try {
    const raw = localStorage.getItem(KEYS.ai);
    if (!raw) return { ...DEFAULT_AI_SETTINGS };
    return { ...DEFAULT_AI_SETTINGS, ...JSON.parse(raw) };
  } catch {
    return { ...DEFAULT_AI_SETTINGS };
  }
}

export function saveAISettings(settings: AISettings): void {
  localStorage.setItem(KEYS.ai, JSON.stringify(settings));
}

export function loadMcpServers(): McpServer[] {
  try {
    const raw = localStorage.getItem(KEYS.mcp);
    if (!raw) return [];
    return JSON.parse(raw);
  } catch {
    return [];
  }
}

export function saveMcpServers(servers: McpServer[]): void {
  localStorage.setItem(KEYS.mcp, JSON.stringify(servers));
}
