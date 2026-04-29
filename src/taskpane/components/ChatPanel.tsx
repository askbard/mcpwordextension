import { useState, useRef, useEffect, useCallback, useContext, KeyboardEvent } from "react";
import {
  Button,
  Textarea,
  makeStyles,
  tokens,
  Text,
  Spinner,
  Badge,
  Tooltip,
} from "@fluentui/react-components";
import {
  SendRegular,
  StopRegular,
  DocumentAddRegular,
  CopyRegular,
  DeleteRegular,
} from "@fluentui/react-icons";
import { McpConnectionsContext } from "./McpConnectionsContext";
import { useAgentLoop, AgentMessage } from "../hooks/useAgentLoop";
import { Message } from "../../providers/index";
import { insertTextAtCursor } from "../../office/bridge";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    overflow: "hidden",
  },
  messages: {
    flex: 1,
    overflowY: "auto",
    padding: "8px 10px",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  bubble: {
    maxWidth: "100%",
    padding: "8px 10px",
    borderRadius: "8px",
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase300,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
  userBubble: {
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground2,
    alignSelf: "flex-end",
    maxWidth: "85%",
  },
  assistantBubble: {
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground1,
    alignSelf: "flex-start",
    width: "100%",
  },
  toolBubble: {
    backgroundColor: tokens.colorNeutralBackground4,
    color: tokens.colorNeutralForeground3,
    alignSelf: "flex-start",
    fontSize: tokens.fontSizeBase100,
    fontFamily: "monospace",
    width: "100%",
  },
  bubbleActions: {
    display: "flex",
    gap: "4px",
    marginTop: "4px",
    opacity: 0,
    transition: "opacity 0.1s",
    "&:hover": { opacity: 1 },
  },
  messageRow: {
    display: "flex",
    flexDirection: "column",
    "&:hover .bubble-actions": { opacity: 1 },
  },
  inputRow: {
    display: "flex",
    gap: "6px",
    padding: "8px 10px",
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
    alignItems: "flex-end",
    flexShrink: 0,
  },
  textarea: {
    flex: 1,
    minHeight: "60px",
    maxHeight: "120px",
    resize: "none",
  },
  toolbar: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    padding: "4px 10px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    flexShrink: 0,
  },
  emptyState: {
    flex: 1,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    color: tokens.colorNeutralForeground4,
    textAlign: "center",
    padding: "20px",
    fontSize: tokens.fontSizeBase200,
  },
});

export default function ChatPanel() {
  const styles = useStyles();
  const { connections } = useContext(McpConnectionsContext);
  const [displayMessages, setDisplayMessages] = useState<AgentMessage[]>([]);
  const [history, setHistory] = useState<Message[]>([]);
  const [input, setInput] = useState("");
  const [isRunning, setIsRunning] = useState(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);

  const onMessage = useCallback((msg: AgentMessage) => {
    setDisplayMessages((prev) => [...prev, msg]);
  }, []);

  const onUpdateLast = useCallback((id: string, delta: string, done?: boolean) => {
    setDisplayMessages((prev) =>
      prev.map((m) =>
        m.id === id
          ? { ...m, content: done && delta === "" ? m.content : m.content + delta, isStreaming: !done }
          : m
      )
    );
  }, []);

  const { run, abort } = useAgentLoop({ connections, onMessage, onUpdateLast });

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [displayMessages]);

  async function handleSend() {
    if (!input.trim() || isRunning) return;

    const userText = input.trim();
    setInput("");

    const userMsg: AgentMessage = {
      id: `user_${Date.now()}`,
      role: "user",
      content: userText,
    };
    setDisplayMessages((prev) => [...prev, userMsg]);

    const newHistory: Message[] = [...history, { role: "user", content: userText }];
    setHistory(newHistory);
    setIsRunning(true);

    try {
      await run(newHistory);
    } finally {
      setIsRunning(false);
    }
  }

  function handleKeyDown(e: KeyboardEvent<HTMLTextAreaElement>) {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  }

  function handleClear() {
    setDisplayMessages([]);
    setHistory([]);
  }

  async function handleInsert(text: string) {
    try {
      await insertTextAtCursor(text);
    } catch (e) {
      console.error("Insert failed:", e);
    }
  }

  function handleCopy(text: string) {
    navigator.clipboard.writeText(text);
  }

  const activeTools = connections.flatMap((c) => c.tools).length;

  return (
    <div className={styles.root}>
      <div className={styles.toolbar}>
        {activeTools > 0 && (
          <Badge color="brand" size="small">
            {activeTools} tool{activeTools !== 1 ? "s" : ""}
          </Badge>
        )}
        <div style={{ flex: 1 }} />
        <Tooltip content="Clear conversation" relationship="label">
          <Button
            appearance="subtle"
            size="small"
            icon={<DeleteRegular />}
            onClick={handleClear}
            disabled={isRunning}
          />
        </Tooltip>
      </div>

      <div className={styles.messages}>
        {displayMessages.length === 0 ? (
          <div className={styles.emptyState}>
            <Text>
              Ask anything. Select text in your document first to give context,
              or connect MCP servers in the MCP tab for tool access.
            </Text>
          </div>
        ) : (
          displayMessages.map((msg) => (
            <div key={msg.id} className={styles.messageRow}>
              <div
                className={`${styles.bubble} ${
                  msg.role === "user"
                    ? styles.userBubble
                    : msg.role === "tool_result"
                    ? styles.toolBubble
                    : styles.assistantBubble
                }`}
              >
                {msg.role === "tool_result" && (
                  <Text size={100} style={{ color: tokens.colorBrandForeground1 }}>
                    🔧 {msg.toolName}
                    {"\n"}
                  </Text>
                )}
                {msg.content}
                {msg.isStreaming && <Spinner size="tiny" style={{ display: "inline-block", marginLeft: 6 }} />}
              </div>
              {msg.role === "assistant" && !msg.isStreaming && msg.content && (
                <div className={`${styles.bubbleActions} bubble-actions`}>
                  <Tooltip content="Insert at cursor" relationship="label">
                    <Button
                      size="small"
                      appearance="subtle"
                      icon={<DocumentAddRegular />}
                      onClick={() => handleInsert(msg.content)}
                    />
                  </Tooltip>
                  <Tooltip content="Copy" relationship="label">
                    <Button
                      size="small"
                      appearance="subtle"
                      icon={<CopyRegular />}
                      onClick={() => handleCopy(msg.content)}
                    />
                  </Tooltip>
                </div>
              )}
            </div>
          ))
        )}
        <div ref={messagesEndRef} />
      </div>

      <div className={styles.inputRow}>
        <Textarea
          className={styles.textarea}
          value={input}
          onChange={(_, d) => setInput(d.value)}
          onKeyDown={handleKeyDown}
          placeholder="Message… (Enter to send, Shift+Enter for newline)"
          disabled={isRunning}
        />
        {isRunning ? (
          <Button appearance="subtle" icon={<StopRegular />} onClick={abort} />
        ) : (
          <Button
            appearance="primary"
            icon={<SendRegular />}
            onClick={handleSend}
            disabled={!input.trim()}
          />
        )}
      </div>
    </div>
  );
}
