import { useState, useEffect } from "react";
import {
  Button,
  Field,
  Input,
  Select,
  Textarea,
  makeStyles,
  tokens,
  Text,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import { SaveRegular } from "@fluentui/react-icons";
import {
  AIProvider,
  AISettings,
  DEFAULT_AI_SETTINGS,
  loadAISettings,
  saveAISettings,
} from "../../storage/settings";

const PROVIDER_LABELS: Record<AIProvider, string> = {
  openrouter: "OpenRouter (Recommended)",
  openai: "OpenAI",
  "openai-compat": "OpenAI-Compatible / Custom",
  google: "Google Gemini",
};

const PROVIDER_BASE_URLS: Record<AIProvider, string> = {
  openrouter: "https://openrouter.ai/api/v1",
  openai: "https://api.openai.com/v1",
  "openai-compat": "",
  google: "https://generativelanguage.googleapis.com/v1beta/openai",
};

const PROVIDER_MODELS: Record<AIProvider, string[]> = {
  openrouter: [
    "anthropic/claude-3.5-sonnet",
    "anthropic/claude-3-opus",
    "anthropic/claude-3-haiku",
    "openai/gpt-4o",
    "openai/gpt-4o-mini",
    "openai/o1",
    "google/gemini-2.0-flash-001",
    "meta-llama/llama-3.3-70b-instruct",
    "mistralai/mistral-large",
  ],
  openai: ["gpt-4o", "gpt-4o-mini", "gpt-4-turbo", "o1", "o1-mini"],
  "openai-compat": [],
  google: ["gemini-2.0-flash", "gemini-1.5-pro", "gemini-1.5-flash"],
};

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "12px",
    overflowY: "auto",
    height: "100%",
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  sectionTitle: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase200,
    textTransform: "uppercase",
    letterSpacing: "0.05em",
  },
  saveRow: {
    display: "flex",
    justifyContent: "flex-end",
    marginTop: "4px",
  },
});

export default function SettingsPanel() {
  const styles = useStyles();
  const [settings, setSettings] = useState<AISettings>(loadAISettings);
  const [saved, setSaved] = useState(false);
  const [customModel, setCustomModel] = useState("");

  const presetModels = PROVIDER_MODELS[settings.provider];
  const isCustomModel = presetModels.length > 0 && !presetModels.includes(settings.model);

  useEffect(() => {
    if (saved) {
      const t = setTimeout(() => setSaved(false), 2000);
      return () => clearTimeout(t);
    }
  }, [saved]);

  function handleProviderChange(provider: AIProvider) {
    setSettings((s) => ({
      ...s,
      provider,
      baseUrl: PROVIDER_BASE_URLS[provider] || s.baseUrl,
      model: PROVIDER_MODELS[provider][0] ?? s.model,
    }));
  }

  function handleSave() {
    const finalSettings = {
      ...settings,
      model: customModel.trim() || settings.model,
    };
    saveAISettings(finalSettings);
    setSettings(finalSettings);
    setCustomModel("");
    setSaved(true);
  }

  function handleReset() {
    setSettings({ ...DEFAULT_AI_SETTINGS });
    setCustomModel("");
  }

  return (
    <div className={styles.root}>
      <div className={styles.section}>
        <Text className={styles.sectionTitle}>AI Provider</Text>

        <Field label="Provider">
          <Select
            value={settings.provider}
            onChange={(_, d) => handleProviderChange(d.value as AIProvider)}
          >
            {(Object.entries(PROVIDER_LABELS) as [AIProvider, string][]).map(([v, label]) => (
              <option key={v} value={v}>{label}</option>
            ))}
          </Select>
        </Field>

        <Field label="API Key">
          <Input
            type="password"
            value={settings.apiKey}
            onChange={(_, d) => setSettings((s) => ({ ...s, apiKey: d.value }))}
            placeholder={
              settings.provider === "openrouter"
                ? "sk-or-..."
                : settings.provider === "openai"
                ? "sk-..."
                : "Your API key"
            }
          />
        </Field>

        <Field label="Base URL">
          <Input
            value={settings.baseUrl}
            onChange={(_, d) => setSettings((s) => ({ ...s, baseUrl: d.value }))}
            placeholder="https://..."
          />
        </Field>

        <Field label="Model">
          {presetModels.length > 0 ? (
            <Select
              value={isCustomModel ? "__custom__" : settings.model}
              onChange={(_, d) => {
                if (d.value === "__custom__") {
                  setCustomModel(settings.model);
                } else {
                  setSettings((s) => ({ ...s, model: d.value }));
                  setCustomModel("");
                }
              }}
            >
              {presetModels.map((m) => (
                <option key={m} value={m}>{m}</option>
              ))}
              <option value="__custom__">Custom model ID...</option>
            </Select>
          ) : null}
          {(presetModels.length === 0 || isCustomModel || customModel) && (
            <Input
              value={customModel || (isCustomModel ? settings.model : "")}
              onChange={(_, d) => setCustomModel(d.value)}
              placeholder="e.g. anthropic/claude-3.5-sonnet"
            />
          )}
        </Field>
      </div>

      <div className={styles.section}>
        <Text className={styles.sectionTitle}>Behavior</Text>
        <Field label="System Prompt">
          <Textarea
            value={settings.systemPrompt}
            onChange={(_, d) => setSettings((s) => ({ ...s, systemPrompt: d.value }))}
            rows={4}
            resize="vertical"
          />
        </Field>
      </div>

      {saved && (
        <MessageBar intent="success">
          <MessageBarBody>Settings saved.</MessageBarBody>
        </MessageBar>
      )}

      <div className={styles.saveRow}>
        <Button appearance="subtle" onClick={handleReset} style={{ marginRight: 8 }}>
          Reset
        </Button>
        <Button appearance="primary" icon={<SaveRegular />} onClick={handleSave}>
          Save
        </Button>
      </div>
    </div>
  );
}
