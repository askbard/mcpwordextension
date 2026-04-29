import React, { createContext, useState, useCallback, ReactNode } from "react";
import { McpConnection } from "../../mcp/client";

type McpConnectionsContextType = {
  connections: McpConnection[];
  setConnections: React.Dispatch<React.SetStateAction<McpConnection[]>>;
  removeConnection: (serverId: string) => void;
};

export const McpConnectionsContext = createContext<McpConnectionsContextType>({
  connections: [],
  setConnections: () => {},
  removeConnection: () => {},
});

export function McpConnectionsProvider({ children }: { children: ReactNode }) {
  const [connections, setConnections] = useState<McpConnection[]>([]);

  const removeConnection = useCallback((serverId: string) => {
    setConnections((prev) => prev.filter((c) => c.server.id !== serverId));
  }, []);

  return (
    <McpConnectionsContext.Provider value={{ connections, setConnections, removeConnection }}>
      {children}
    </McpConnectionsContext.Provider>
  );
}
