import { useState, useContext } from "react";
import {
  Button,
  Field,
  Input,
  Switch,
  makeStyles,
  tokens,
  Text,
  Badge,
  Spinner,
  MessageBar,
  MessageBarBody,
  Tooltip,
} from "@fluentui/react-components";
import {
  AddRegular,
  DeleteRegular,
  PlugConnectedRegular,
  PlugDisconnectedRegular,
  ArrowSyncRegular,
} from "@fluentui/react-icons";
import { McpServer, loadMcpServers, saveMcpServers } from "../../storage/settings";
import { connectMcpServer, disconnectMcpServer, McpConnection } from "../../mcp/client";
import { McpConnectionsContext } from "./McpConnectionsContext";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    padding: "10px",
    overflowY: "auto",
    height: "100%",
  },
  serverRow: {
    display: "flex",
    flexDirection: "column",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: "6px",
    overflow: "hidden",
  },
  serverHeader: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    padding: "8px 10px",
    backgroundColor: tokens.colorNeutralBackground2,
  },
  serverName: {
    flex: 1,
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  serverUrl: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    padding: "0 10px 6px",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  toolList: {
    padding: "4px 10px 8px",
    display: "flex",
    flexDirection: "column",
    gap: "2px",
  },
  toolItem: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground2,
    padding: "2px 4px",
    borderRadius: "4px",
    backgroundColor: tokens.colorNeutralBackground3,
  },
  addForm: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "10px",
    border: `1px solid ${tokens.colorBrandStroke1}`,
    borderRadius: "6px",
    backgroundColor: tokens.colorBrandBackground2,
  },
  formTitle: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
  },
  headerRow: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    padding: "2px 0",
  },
  sectionLabel: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase100,
    textTransform: "uppercase",
    letterSpacing: "0.05em",
  },
});

type ServerStatus = "idle" | "connecting" | "connected" | "error";

export default function McpPanel() {
  const styles = useStyles();
  const { connections, setConnections, removeConnection } = useContext(McpConnectionsContext);
  const [servers, setServers] = useState<McpServer[]>(loadMcpServers);
  const [statuses, setStatuses] = useState<Record<string, ServerStatus>>({});
  const [errors, setErrors] = useState<Record<string, string>>({});
  const [showAddForm, setShowAddForm] = useState(false);

  const [newName, setNewName] = useState("");
  const [newUrl, setNewUrl] = useState("");
  const [newHeaders, setNewHeaders] = useState("");

  function saveServers(updated: McpServer[]) {
    setServers(updated);
    saveMcpServers(updated);
  }

  function handleAdd() {
    if (!newName.trim() || !newUrl.trim()) return;

    let headers: Record<string, string> = {};
    if (newHeaders.trim()) {
      try {
        headers = JSON.parse(newHeaders);
      } catch {
        // treat as key: value lines
        for (const line of newHeaders.split("\n")) {
          const [k, ...rest] = line.split(":");
          if (k && rest.length) headers[k.trim()] = rest.join(":").trim();
        }
      }
    }

    const server: McpServer = {
      id: `mcp_${Date.now()}`,
      name: newName.trim(),
      url: newUrl.trim(),
      headers,
      enabled: true,
    };

    saveServers([...servers, server]);
    setNewName("");
    setNewUrl("");
    setNewHeaders("");
    setShowAddForm(false);
  }

  function handleDelete(id: string) {
    const conn = connections.find((c) => c.server.id === id);
    if (conn) {
      disconnectMcpServer(conn).catch(console.error);
      removeConnection(id);
    }
    saveServers(servers.filter((s) => s.id !== id));
    setStatuses((p) => { const n = { ...p }; delete n[id]; return n; });
    setErrors((p) => { const n = { ...p }; delete n[id]; return n; });
  }

  function handleToggleEnabled(id: string, enabled: boolean) {
    saveServers(servers.map((s) => s.id === id ? { ...s, enabled } : s));
    if (!enabled) {
      const conn = connections.find((c) => c.server.id === id);
      if (conn) {
        disconnectMcpServer(conn).catch(console.error);
        removeConnection(id);
      }
      setStatuses((p) => ({ ...p, [id]: "idle" }));
    }
  }

  async function handleConnect(server: McpServer) {
    setStatuses((p) => ({ ...p, [server.id]: "connecting" }));
    setErrors((p) => { const n = { ...p }; delete n[server.id]; return n; });

    // Disconnect existing if any
    const existing = connections.find((c) => c.server.id === server.id);
    if (existing) {
      await disconnectMcpServer(existing).catch(console.error);
      removeConnection(server.id);
    }

    try {
      const conn = await connectMcpServer(server);
      setConnections((prev) => [...prev.filter((c) => c.server.id !== server.id), conn]);
      setStatuses((p) => ({ ...p, [server.id]: "connected" }));
    } catch (e) {
      setStatuses((p) => ({ ...p, [server.id]: "error" }));
      setErrors((p) => ({ ...p, [server.id]: e instanceof Error ? e.message : String(e) }));
    }
  }

  function getConnection(id: string): McpConnection | undefined {
    return connections.find((c) => c.server.id === id);
  }

  return (
    <div className={styles.root}>
      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
        <Text className={styles.sectionLabel}>MCP Servers</Text>
        <div style={{ flex: 1 }} />
        <Button
          size="small"
          appearance="primary"
          icon={<AddRegular />}
          onClick={() => setShowAddForm((v) => !v)}
        >
          Add
        </Button>
      </div>

      {showAddForm && (
        <div className={styles.addForm}>
          <Text className={styles.formTitle}>New MCP Server</Text>
          <Field label="Name">
            <Input
              value={newName}
              onChange={(_, d) => setNewName(d.value)}
              placeholder="e.g. Web Search"
              size="small"
            />
          </Field>
          <Field label="URL">
            <Input
              value={newUrl}
              onChange={(_, d) => setNewUrl(d.value)}
              placeholder="https://your-mcp-server.com/mcp"
              size="small"
            />
          </Field>
          <Field
            label="Auth Headers (optional)"
            hint='JSON object or "Key: Value" lines'
          >
            <Input
              value={newHeaders}
              onChange={(_, d) => setNewHeaders(d.value)}
              placeholder='{"Authorization": "Bearer token"}'
              size="small"
            />
          </Field>
          <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
            <Button size="small" appearance="subtle" onClick={() => setShowAddForm(false)}>
              Cancel
            </Button>
            <Button
              size="small"
              appearance="primary"
              onClick={handleAdd}
              disabled={!newName.trim() || !newUrl.trim()}
            >
              Add Server
            </Button>
          </div>
        </div>
      )}

      {servers.length === 0 && (
        <Text style={{ color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase100 }}>
          No MCP servers configured. Add one to give the AI access to external tools.
        </Text>
      )}

      {servers.map((server) => {
        const status = statuses[server.id] ?? "idle";
        const conn = getConnection(server.id);
        const err = errors[server.id];

        return (
          <div key={server.id} className={styles.serverRow}>
            <div className={styles.serverHeader}>
              {status === "connecting" && <Spinner size="extra-tiny" />}
              {status === "connected" && <PlugConnectedRegular style={{ color: tokens.colorPaletteGreenForeground1 }} />}
              {(status === "idle" || status === "error") && <PlugDisconnectedRegular style={{ color: tokens.colorNeutralForeground3 }} />}

              <Text className={styles.serverName}>{server.name}</Text>

              {conn && (
                <Badge color="success" size="small">
                  {conn.tools.length} tool{conn.tools.length !== 1 ? "s" : ""}
                </Badge>
              )}

              <Switch
                checked={server.enabled}
                onChange={(_, d) => handleToggleEnabled(server.id, d.checked)}
              />

              <Tooltip content="Connect / Reconnect" relationship="label">
                <Button
                  size="small"
                  appearance="subtle"
                  icon={<ArrowSyncRegular />}
                  onClick={() => handleConnect(server)}
                  disabled={status === "connecting" || !server.enabled}
                />
              </Tooltip>

              <Tooltip content="Delete" relationship="label">
                <Button
                  size="small"
                  appearance="subtle"
                  icon={<DeleteRegular />}
                  onClick={() => handleDelete(server.id)}
                />
              </Tooltip>
            </div>

            <Text className={styles.serverUrl}>{server.url}</Text>

            {err && (
              <MessageBar intent="error" style={{ margin: "0 8px 6px" }}>
                <MessageBarBody>{err}</MessageBarBody>
              </MessageBar>
            )}

            {conn && conn.tools.length > 0 && (
              <div className={styles.toolList}>
                {conn.tools.map((t) => (
                  <Text key={t.name} className={styles.toolItem}>
                    {t.name.replace(`${server.name}__`, "")}
                  </Text>
                ))}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
