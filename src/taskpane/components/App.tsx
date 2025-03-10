import * as React from "react";
import { DefaultButton, MessageBar, MessageBarType, ProgressIndicator, TextField } from "@fluentui/react";
import axios from "axios"; // 修改为 axios
import Center from "./Center";
import Container from "./Container";
import Login from "./Login";
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
          model: "deepseek-chat",
          messages: [{ role: "user", content: prompt }], // 使用messages数组
          max_tokens: 2048,
          temperature: 0.7,
        },
        {
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${apiKey}`,
          },
          timeout: 30000,
        }
      );

      // 修正响应数据路径
      setGeneratedText(response.data.choices[0].message.content);
      setLoading(false);
      setError("");
    } catch (error: any) {
      if (error.response) {
        const status = error.response.status;
        setError(`Error: ${status} - ${error.response.data?.message || "Unknown error"}`);
        // 仅在授权失败时清除API密钥
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
            rows={5}
            multiline={true}
            onChange={(_, newValue?: string) => setPrompt(newValue || "")}
          ></TextField>
          <Center>
            <DefaultButton iconProps={{ iconName: "Robot" }} onClick={onClick}>
              Generate
            </DefaultButton>
          </Center>
          {loading && <ProgressIndicator label="Generating text..." />}
          {generatedText && (
            <div>
              <p className="generated-text">{generatedText}</p>
              <Center>
                <DefaultButton iconProps={{ iconName: "Add" }} onClick={onInsert}>
                  Insert text
                </DefaultButton>
                <DefaultButton iconProps={{ iconName: "Copy" }} onClick={onCopy}>
                  Copy text
                </DefaultButton>
              </Center>
            </div>
          )}
        </>
      ) : (
        <Login onSave={saveApiKey} />
      )}
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
    </Container>
  );
}
