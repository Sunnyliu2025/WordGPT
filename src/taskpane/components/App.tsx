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
            placeholder="Enter prompt here"
            value={prompt}
            rows={8}
            multiline={true}
            onChange={(_, newValue?: string) => setPrompt(newValue || "")}
          />
          <Center>
            <CommandBarButton
              iconProps={{ iconName: "Send" }}
              onClick={onClick}
              styles={{
                root: {
                  backgroundColor: "#0078d4",
                  color: "white",
                  margin: "12px 0",
                  borderRadius: "6px",
                  padding: "0 28px",
                  minWidth: "140px",
                  height: "40px",
                  border: "none",
                  boxShadow: "0 2px 8px rgba(0, 120, 212, 0.3)",
                  transition: "all 0.25s ease",
                  selectors: {
                    ":hover": {
                      backgroundColor: "#106ebe",
                      boxShadow: "0 4px 12px rgba(0, 120, 212, 0.4)",
                    },
                    ":active": {
                      backgroundColor: "#005a9e",
                    },
                    ":disabled": {
                      backgroundColor: "#ccc",
                      cursor: "not-allowed",
                      boxShadow: "none",
                    },
                  },
                },
                icon: {
                  color: "white",
                  fontSize: "16px",
                },
                label: {
                  fontWeight: 600,
                  fontSize: "14px",
                  textTransform: "uppercase",
                  letterSpacing: "0.5px",
                },
                flexContainer: {
                  justifyContent: "center",
                },
              }}
            >
              Generate
            </CommandBarButton>
          </Center>
          {loading && (
            <div className="loading-container">
              <ProgressIndicator label="Generating text..." />
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
