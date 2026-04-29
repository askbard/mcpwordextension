import { useState } from "react";
import {
  Button,
  Input,
  makeStyles,
  tokens,
  Text,
  Spinner,
  Divider,
  MessageBar,
  MessageBarBody,
  Tooltip,
} from "@fluentui/react-components";
import {
  DocumentRegular,
  WandRegular,
  TranslateRegular,
  TextGrammarErrorRegular,
  ArrowExpandRegular,
  ArrowCollapseAllRegular,
  TextEditStyleRegular,
  DocumentAddRegular,
  ArrowSwapRegular,
  CopyRegular,
  ArrowClockwiseRegular,
} from "@fluentui/react-icons";
import { getSelectedText, insertTextAtCursor, insertTextBelow } from "../../office/bridge";
import { streamChat } from "../../providers/index";
import { loadAISettings } from "../../storage/settings";

type Action = {
  label: string;
  icon: React.ReactElement;
  prompt: (selected: string) => string;
};

const ACTIONS: Action[] = [
  {
    label: "Summarize",
    icon: <DocumentRegular />,
    prompt: (s) => `Summarize the following text concisely:\n\n${s}`,
  },
  {
    label: "Rewrite",
    icon: <TextEditStyleRegular />,
    prompt: (s) => `Rewrite the following text to improve clarity and flow, keeping the same meaning:\n\n${s}`,
  },
  {
    label: "Fix Grammar",
    icon: <TextGrammarErrorRegular />,
    prompt: (s) => `Fix all grammar, spelling, and punctuation errors in the following text. Return only the corrected text:\n\n${s}`,
  },
  {
    label: "Expand",
    icon: <ArrowExpandRegular />,
    prompt: (s) => `Expand the following text with more detail and explanation:\n\n${s}`,
  },
  {
    label: "Shorten",
    icon: <ArrowCollapseAllRegular />,
    prompt: (s) => `Shorten the following text while keeping the key points:\n\n${s}`,
  },
  {
    label: "Translate",
    icon: <TranslateRegular />,
    prompt: (s) => `Translate the following text to English (or if already in English, translate to Spanish):\n\n${s}`,
  },
  {
    label: "Improve",
    icon: <WandRegular />,
    prompt: (s) => `Improve the writing quality of the following text, making it more professional and engaging:\n\n${s}`,
  },
];

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    padding: "10px",
    height: "100%",
    overflowY: "auto",
  },
  selectionBox: {
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: "6px",
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground2,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    minHeight: "60px",
    maxHeight: "120px",
    overflowY: "auto",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
  noSelection: {
    color: tokens.colorNeutralForeground4,
    fontStyle: "italic",
  },
  actions: {
    display: "flex",
    flexWrap: "wrap",
    gap: "6px",
  },
  result: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    flex: 1,
  },
  resultBox: {
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: "6px",
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground1,
    fontSize: tokens.fontSizeBase200,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    flex: 1,
    minHeight: "80px",
    overflowY: "auto",
  },
  resultActions: {
    display: "flex",
    gap: "6px",
    flexWrap: "wrap",
  },
  sectionLabel: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase100,
    textTransform: "uppercase",
    letterSpacing: "0.05em",
  },
  customRow: {
    display: "flex",
    gap: "6px",
  },
});

export default function ActionsPanel() {
  const styles = useStyles();
  const [selectedText, setSelectedText] = useState<string | null>(null);
  const [loadingSelection, setLoadingSelection] = useState(false);
  const [result, setResult] = useState("");
  const [isRunning, setIsRunning] = useState(false);
  const [customPrompt, setCustomPrompt] = useState("");
  const [error, setError] = useState("");

  async function handleReadSelection() {
    setLoadingSelection(true);
    setError("");
    try {
      const text = await getSelectedText();
      setSelectedText(text || "");
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoadingSelection(false);
    }
  }

  async function runAction(prompt: string) {
    if (!selectedText && !customPrompt) return;
    setResult("");
    setIsRunning(true);
    setError("");

    const settings = loadAISettings();
    if (!settings.apiKey) {
      setError("No API key configured. Go to the Settings tab.");
      setIsRunning(false);
      return;
    }

    try {
      let accumulated = "";
      for await (const chunk of streamChat(
        settings,
        [{ role: "user", content: prompt }]
      )) {
        if (chunk.type === "text") {
          accumulated += chunk.delta;
          setResult(accumulated);
        } else if (chunk.type === "done") {
          break;
        }
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setIsRunning(false);
    }
  }

  async function handleInsert() {
    try {
      await insertTextAtCursor(result);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }

  async function handleInsertBelow() {
    try {
      await insertTextBelow(result);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }

  return (
    <div className={styles.root}>
      {/* Selection */}
      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
        <Text className={styles.sectionLabel}>Selected Text</Text>
        <div style={{ flex: 1 }} />
        <Button
          size="small"
          appearance="subtle"
          icon={loadingSelection ? <Spinner size="extra-tiny" /> : <ArrowClockwiseRegular />}
          onClick={handleReadSelection}
          disabled={loadingSelection}
        >
          Read Selection
        </Button>
      </div>

      <div className={styles.selectionBox}>
        {selectedText === null ? (
          <span className={styles.noSelection}>Click "Read Selection" to load selected text from your document.</span>
        ) : selectedText === "" ? (
          <span className={styles.noSelection}>No text selected in document.</span>
        ) : (
          selectedText
        )}
      </div>

      <Divider />

      {/* Quick actions */}
      <Text className={styles.sectionLabel}>Quick Actions</Text>
      <div className={styles.actions}>
        {ACTIONS.map((action) => (
          <Button
            key={action.label}
            size="small"
            appearance="secondary"
            icon={action.icon}
            onClick={() => runAction(action.prompt(selectedText ?? ""))}
            disabled={!selectedText || isRunning}
          >
            {action.label}
          </Button>
        ))}
      </div>

      {/* Custom prompt */}
      <Text className={styles.sectionLabel}>Custom Prompt</Text>
      <div className={styles.customRow}>
        <Input
          style={{ flex: 1 }}
          size="small"
          value={customPrompt}
          onChange={(_, d) => setCustomPrompt(d.value)}
          placeholder="e.g. Convert to bullet points"
          onKeyDown={(e) => {
            if (e.key === "Enter" && customPrompt.trim()) {
              const fullPrompt = selectedText
                ? `${customPrompt}\n\n${selectedText}`
                : customPrompt;
              runAction(fullPrompt);
            }
          }}
        />
        <Button
          size="small"
          appearance="primary"
          icon={isRunning ? <Spinner size="extra-tiny" /> : <WandRegular />}
          onClick={() => {
            const fullPrompt = selectedText
              ? `${customPrompt}\n\n${selectedText}`
              : customPrompt;
            runAction(fullPrompt);
          }}
          disabled={!customPrompt.trim() || isRunning}
        >
          Run
        </Button>
      </div>

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      {/* Result */}
      {(result || isRunning) && (
        <>
          <Divider />
          <Text className={styles.sectionLabel}>Result</Text>
          <div className={styles.resultBox}>
            {result}
            {isRunning && <Spinner size="tiny" style={{ display: "inline-block", marginLeft: 4 }} />}
          </div>
          {result && !isRunning && (
            <div className={styles.resultActions}>
              <Tooltip content="Replace selection" relationship="label">
                <Button
                  size="small"
                  appearance="primary"
                  icon={<ArrowSwapRegular />}
                  onClick={handleInsert}
                >
                  Replace
                </Button>
              </Tooltip>
              <Tooltip content="Insert below selection" relationship="label">
                <Button
                  size="small"
                  appearance="secondary"
                  icon={<DocumentAddRegular />}
                  onClick={handleInsertBelow}
                >
                  Insert Below
                </Button>
              </Tooltip>
              <Button
                size="small"
                appearance="subtle"
                icon={<CopyRegular />}
                onClick={() => navigator.clipboard.writeText(result)}
              >
                Copy
              </Button>
            </div>
          )}
        </>
      )}
    </div>
  );
}
