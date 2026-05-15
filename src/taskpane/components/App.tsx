import * as React from "react";
import { CommandButton, MessageBar, MessageBarType, ProgressIndicator, TextField } from "@fluentui/react";
import { CommandBarButton } from "@fluentui/react/lib/Button";
import axios from "axios";
import Center from "./Center";
import Container from "./Container";
import Login from "./Login";
import "./initializeIcons";
/* global Word, localStorage, navigator */

export default function App() {
  const [apiKey, setApiKey] = React.useState<string>("");
  const [prompt, setPrompt] = React.useState<string>("");
  const [error, setError] = React.useState<string>("");
  const [loading, setLoading] = React.useState<boolean>(false);
  const [generatedText, setGeneratedText] = React.useState<string>("");

  React.useEffect(() => {
    const key = localStorage.getItem("apiKey");
    if (key) {
      setApiKey(key);
    }
  }, []);

  const saveApiKey = (key: string) => {
    setApiKey(key);
    localStorage.setItem("apiKey", key);
    setError("");
  };

  const onClick = async () => {
    setGeneratedText("");
    setLoading(true);

    try {
      const response = await axios.post(
        "https://api.deepseek.com/v1/chat/completions",
        {
          model: "deepseek-v4-flash",
          messages: [{ role: "user", content: prompt }],
          max_tokens: 8192,
          temperature: 0.7,
        },
        {
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${apiKey}`,
          },
          timeout: 60000,
        }
      );

      setGeneratedText(response.data.choices[0].message.content);
      setLoading(false);
      setError("");
    } catch (error: any) {
      if (error.response) {
        const status = error.response.status;
        setError(`Error: ${status} - ${error.response.data?.message || "Unknown error"}`);
        if (status === 401) {
          setApiKey("");
          localStorage.removeItem("apiKey");
        }
      } else if (error.request) {
        setError("Error: No response received from server.");
      } else {
        setError(`Error: ${error.message}`);
      }
      setLoading(false);
    }
  };

  const onInsert = async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(generatedText, "Start");
      await context.sync();
    });
  };

  const onCopy = async () => {
    navigator.clipboard.writeText(generatedText);
  };

  return (
    <Container>
      {apiKey ? (
        <>
          <TextField
            placeholder="输入提示词"
            value={prompt}
            rows={12}
            multiline={true}
            onChange={(_, newValue?: string) => setPrompt(newValue || "")}
          />
          <Center>
            <CommandBarButton
              iconProps={{ iconName: "Cloud" }}
              onClick={onClick}
              styles={{
                root: {
                  background: "linear-gradient(135deg, #0a7e3c 0%, #1aad5a 50%, #0a7e3c 100%)",
                  color: "white",
                  margin: "28px 0 16px",
                  borderRadius: "14px",
                  padding: "0 52px",
                  minWidth: "240px",
                  height: "60px",
                  border: "none",
                  position: "relative",
                  overflow: "hidden",
                  boxShadow: "0 4px 16px rgba(10, 126, 60, 0.4)",
                  transition: "all 0.3s cubic-bezier(0.4, 0, 0.2, 1)",
                  selectors: {
                    ":hover": {
                      background: "linear-gradient(135deg, #0b8a42 0%, #1fbf62 50%, #0b8a42 100%)",
                      boxShadow: "0 6px 24px rgba(10, 126, 60, 0.55)",
                      transform: "translateY(-2px) scale(1.02)",
                    },
                    ":active": {
                      background: "linear-gradient(135deg, #086b32 0%, #15984d 50%, #086b32 100%)",
                      boxShadow: "0 2px 8px rgba(10, 126, 60, 0.4)",
                      transform: "translateY(0) scale(0.98)",
                    },
                    ":disabled": {
                      background: "#ccc",
                      cursor: "not-allowed",
                      boxShadow: "none",
                      transform: "none",
                    },
                  },
                },
                icon: {
                  color: "white",
                  fontSize: "20px",
                  marginRight: "8px",
                },
                label: {
                  fontWeight: 700,
                  fontSize: "18px",
                  textTransform: "uppercase",
                  letterSpacing: "1.2px",
                },
                flexContainer: {
                  justifyContent: "center",
                  alignItems: "center",
                },
              }}
            >
              生成
            </CommandBarButton>
          </Center>
          {loading && (
            <div className="loading-container">
              <ProgressIndicator label="生成中..." />
            </div>
          )}
          {generatedText && (
            <div>
              <div className="generated-text">{generatedText}</div>
              <div className="button-group">
                <CommandButton className="btn-action" iconProps={{ iconName: "AddTo" }} onClick={onInsert}>
                  Insert text
                </CommandButton>
                <CommandButton className="btn-action" iconProps={{ iconName: "Copy" }} onClick={onCopy}>
                  Copy text
                </CommandButton>
              </div>
            </div>
          )}
        </>
      ) : (
        <Login onSave={saveApiKey} />
      )}
      {error && (
        <div className="error-message">
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        </div>
      )}
    </Container>
  );
}
