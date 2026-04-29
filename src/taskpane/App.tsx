import { useState } from "react";
import {
  Tab,
  TabList,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  ChatMultipleRegular,
  WandRegular,
  PlugConnectedRegular,
  SettingsRegular,
} from "@fluentui/react-icons";
import ChatPanel from "./components/ChatPanel";
import ActionsPanel from "./components/ActionsPanel";
import McpPanel from "./components/McpPanel";
import SettingsPanel from "./components/SettingsPanel";
import { McpConnectionsProvider } from "./components/McpConnectionsContext";

type TabId = "chat" | "actions" | "mcp" | "settings";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    overflow: "hidden",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  tabList: {
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    padding: "4px 8px 0",
    flexShrink: 0,
  },
  content: {
    flex: 1,
    overflow: "hidden",
    display: "flex",
    flexDirection: "column",
  },
});

export default function App() {
  const styles = useStyles();
  const [activeTab, setActiveTab] = useState<TabId>("chat");

  return (
    <McpConnectionsProvider>
      <div className={styles.root}>
        <TabList
          className={styles.tabList}
          selectedValue={activeTab}
          onTabSelect={(_, d) => setActiveTab(d.value as TabId)}
          size="small"
        >
          <Tab value="chat" icon={<ChatMultipleRegular />}>Chat</Tab>
          <Tab value="actions" icon={<WandRegular />}>Actions</Tab>
          <Tab value="mcp" icon={<PlugConnectedRegular />}>MCP</Tab>
          <Tab value="settings" icon={<SettingsRegular />}>Settings</Tab>
        </TabList>

        <div className={styles.content}>
          {activeTab === "chat" && <ChatPanel />}
          {activeTab === "actions" && <ActionsPanel />}
          {activeTab === "mcp" && <McpPanel />}
          {activeTab === "settings" && <SettingsPanel />}
        </div>
      </div>
    </McpConnectionsProvider>
  );
}
