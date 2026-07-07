import * as React from "react";
import { CommandButton, DefaultButton, IconButton, MessageBar, MessageBarType, TextField } from "@fluentui/react";
import axios from "axios";
import Center from "./Center";
import Container from "./Container";
import Login from "./Login";
import "./initializeIcons";
/* global Word, localStorage, navigator, setInterval, clearInterval, setTimeout */

export default function App() {
  const [apiKey, setApiKey] = React.useState<string>("");
  const [prompt, setPrompt] = React.useState<string>("");
  const [error, setError] = React.useState<string>("");
  const [loading, setLoading] = React.useState<boolean>(false);
  const [generatedText, setGeneratedText] = React.useState<string>("");

  // 动画状态
  const [showResult, setShowResult] = React.useState<boolean>(false);
  const [dots, setDots] = React.useState<string>("");

  React.useEffect(() => {
    const key = localStorage.getItem("apiKey");
    if (key) {
      setApiKey(key);
    }
  }, []);

  // 加载动画点
  React.useEffect(() => {
    if (loading) {
      const interval = setInterval(() => {
        setDots((prev) => (prev.length < 3 ? prev + "." : ""));
      }, 400);
      return () => clearInterval(interval);
    }
    return undefined;
  }, [loading]);

  const saveApiKey = (key: string) => {
    setApiKey(key);
    localStorage.setItem("apiKey", key);
    setError("");
  };

  const onClick = async () => {
    if (!prompt.trim()) {
      setError("请输入提示词");
      return;
    }
    setGeneratedText("");
    setShowResult(false);
    setLoading(true);
    setError("");

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
      // 延迟显示结果以触发动画
      setTimeout(() => setShowResult(true), 50);
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

  const onClear = () => {
    setPrompt("");
    setGeneratedText("");
    setShowResult(false);
    setError("");
  };

  return (
    <Container>
      <div className="app-wrapper">
        {/* ===== 头部区域 ===== */}
        <header className="app-header">
          <div className="app-header-content">
            <img
              src="assets/deepseeklogo.png"
              alt="DeepSeek"
              className="app-logo"
              onError={(e) => {
                (e.target as HTMLImageElement).style.display = "none";
              }}
            />
            <div className="app-header-text">
              <h1 className="app-title">WordGPT</h1>
              <p className="app-subtitle">DeepSeek AI 写作助手</p>
            </div>
          </div>
          <IconButton
            iconProps={{ iconName: "Clear" }}
            title="清空"
            ariaLabel="清空"
            onClick={onClear}
            styles={{
              root: {
                color: "#6b7280",
                width: 32,
                height: 32,
                borderRadius: 8,
                transition: "all 0.2s ease",
              },
              rootHovered: {
                color: "#374151",
                background: "rgba(0,0,0,0.06)",
              },
              icon: { fontSize: 16, fontWeight: 700 },
            }}
          />
        </header>

        {apiKey ? (
          <>
            {/* ===== 提示词输入区 ===== */}
            <div className="input-section">
              <div className="input-label">
                <span>提示词 (Prompt)</span>
                <span className="char-count">{prompt.length} / 4000</span>
              </div>
              <TextField
                placeholder="输入你的提示词，例如：帮我写一篇关于人工智能的文章..."
                value={prompt}
                rows={10}
                multiline={true}
                resizable={false}
                maxLength={4000}
                onChange={(_, newValue?: string) => setPrompt(newValue || "")}
                styles={{
                  root: { width: "100%" },
                  fieldGroup: {
                    borderRadius: 12,
                    border: "2px solid #e5e7eb",
                    transition: "border-color 0.2s ease, box-shadow 0.2s ease",
                    background: "#f9fafb",
                    selectors: {
                      ":hover": {
                        borderColor: "#3b82f6",
                      },
                      ":focus-within": {
                        borderColor: "#3b82f6",
                        boxShadow: "0 0 0 3px rgba(59, 130, 246, 0.15)",
                        background: "#fff",
                      },
                    },
                  },
                  field: {
                    fontSize: 14,
                    lineHeight: 1.7,
                    color: "#1f2937",
                    padding: "14px 16px",
                    "::placeholder": {
                      color: "#9ca3af",
                      fontSize: 14,
                    },
                  },
                }}
              />
            </div>

            {/* ===== 生成按钮 ===== */}
            <Center>
              <DefaultButton
                iconProps={{ iconName: "Cloud" }}
                onClick={onClick}
                disabled={loading || !prompt.trim()}
                styles={{
                  root: {
                    background: "linear-gradient(135deg, #0a7e3c 0%, #1aad5a 50%, #0a7e3c 100%)",
                    color: "white",
                    margin: "24px 0 20px",
                    borderRadius: 14,
                    padding: "0 48px",
                    minWidth: 220,
                    height: 52,
                    border: "none",
                    position: "relative",
                    overflow: "hidden",
                    boxShadow: "0 4px 16px rgba(10, 126, 60, 0.35)",
                    transition: "all 0.3s cubic-bezier(0.4, 0, 0.2, 1)",
                    cursor: loading || !prompt.trim() ? "not-allowed" : "pointer",
                    opacity: loading || !prompt.trim() ? 0.6 : 1,
                    selectors: {
                      ":hover": {
                        background: "linear-gradient(135deg, #0b8a42 0%, #1fbf62 50%, #0b8a42 100%)",
                        boxShadow: "0 6px 24px rgba(10, 126, 60, 0.5)",
                        transform: loading || !prompt.trim() ? "none" : "translateY(-2px) scale(1.02)",
                      },
                      ":active": {
                        background: "linear-gradient(135deg, #086b32 0%, #15984d 50%, #086b32 100%)",
                        boxShadow: "0 2px 8px rgba(10, 126, 60, 0.4)",
                        transform: loading || !prompt.trim() ? "none" : "translateY(0) scale(0.98)",
                      },
                      "::after": {
                        content: '""',
                        position: "absolute",
                        top: 0,
                        left: 0,
                        right: 0,
                        bottom: 0,
                        background:
                          "linear-gradient(135deg, transparent 0%, rgba(255,255,255,0.12) 50%, transparent 100%)",
                        pointerEvents: "none",
                      },
                    },
                  },
                  icon: {
                    color: "white",
                    fontSize: 18,
                    marginRight: 8,
                  },
                  label: {
                    fontWeight: 700,
                    fontSize: 16,
                    letterSpacing: "1px",
                  },
                  flexContainer: {
                    justifyContent: "center",
                    alignItems: "center",
                  },
                }}
              >
                {loading ? "生成中..." : "✨ 生成"}
              </DefaultButton>
            </Center>

            {/* ===== 加载状态 ===== */}
            {loading && (
              <div className="loading-container">
                <div className="loading-bar">
                  <div className="loading-bar-inner" />
                </div>
                <div className="loading-text">
                  <span className="loading-icon">🧠</span>
                  AI 正在思考{dots}
                </div>
              </div>
            )}

            {/* ===== 生成结果 ===== */}
            {generatedText && (
              <div className={`result-section ${showResult ? "result-visible" : ""}`}>
                <div className="result-header">
                  <span className="result-header-icon">📝</span>
                  <span className="result-header-title">生成结果</span>
                </div>
                <div className="result-content">
                  <p className="result-text">{generatedText}</p>
                </div>
                <div className="result-actions">
                  <CommandButton
                    className="btn-action"
                    iconProps={{ iconName: "AddTo" }}
                    onClick={onInsert}
                    styles={{
                      root: {
                        borderRadius: 8,
                        height: 36,
                        padding: "0 16px",
                        transition: "all 0.2s ease",
                      },
                      rootHovered: {
                        background: "rgba(10, 126, 60, 0.08)",
                      },
                      icon: { color: "#0a7e3c", fontSize: 14 },
                      label: { fontWeight: 600, fontSize: 13, color: "#0a7e3c" },
                    }}
                  >
                    插入文档
                  </CommandButton>
                  <CommandButton
                    className="btn-action"
                    iconProps={{ iconName: "Copy" }}
                    onClick={onCopy}
                    styles={{
                      root: {
                        borderRadius: 8,
                        height: 36,
                        padding: "0 16px",
                        transition: "all 0.2s ease",
                      },
                      rootHovered: {
                        background: "rgba(59, 130, 246, 0.08)",
                      },
                      icon: { color: "#3b82f6", fontSize: 14 },
                      label: { fontWeight: 600, fontSize: 13, color: "#3b82f6" },
                    }}
                  >
                    复制文本
                  </CommandButton>
                </div>
              </div>
            )}
          </>
        ) : (
          <Login onSave={saveApiKey} />
        )}

        {/* ===== 错误消息 ===== */}
        {error && (
          <div className="error-section">
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={false}
              onDismiss={() => setError("")}
              dismissIconProps={{ iconName: "Cancel" }}
            >
              {error}
            </MessageBar>
          </div>
        )}

        {/* ===== 页脚 ===== */}
        <footer className="app-footer">
          <p>Powered by DeepSeek AI</p>
        </footer>
      </div>
    </Container>
  );
}
