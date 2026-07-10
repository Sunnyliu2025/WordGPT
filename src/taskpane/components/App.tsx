import * as React from "react";
import { CommandButton, DefaultButton, IconButton, MessageBar, MessageBarType, TextField } from "@fluentui/react";
import axios from "axios";
import Center from "./Center";
import Container from "./Container";
import Login from "./Login";
/* global Word, localStorage, navigator, console, setInterval, clearInterval, setTimeout */

const MAX_PROMPT_LENGTH = 4000;
const DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions";
const DEEPSEEK_MODEL = "deepseek-v4-flash";

interface ErrorResponse {
  message?: string;
}

export default function App() {
  const [apiKey, setApiKey] = React.useState<string>("");
  const [prompt, setPrompt] = React.useState<string>("");
  const [error, setError] = React.useState<string>("");
  const [loading, setLoading] = React.useState<boolean>(false);
  const [generatedText, setGeneratedText] = React.useState<string>("");

  // 动画状态
  const [showResult, setShowResult] = React.useState<boolean>(false);
  const [dots, setDots] = React.useState<string>("");

  // 延迟初始化 Fluent UI 图标，避免阻塞首次渲染
  // initializeIcons() 内部已处理去重逻辑，多次调用安全
  React.useEffect(() => {
    setTimeout(() => {
      import("./initializeIcons");
    }, 0);
  }, []);

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

  /** 统一处理 API 错误 */
  const handleApiError = (error: unknown) => {
    if (error && typeof error === "object" && "response" in error) {
      // Axios 错误
      const axiosErr = error as { response?: { status?: number; data?: ErrorResponse }; message?: string };
      const status = axiosErr.response?.status;
      const data = axiosErr.response?.data;
      setError(`错误 ${status || "未知"}: ${data?.message || axiosErr.message || "未知错误"}`);
      if (status === 401) {
        setApiKey("");
        localStorage.removeItem("apiKey");
      }
    } else if (error instanceof Error) {
      setError(`错误: ${error.message}`);
    } else {
      setError("发生未知错误");
    }
  };

  const onClick = async () => {
    const trimmedPrompt = prompt.trim();
    if (!trimmedPrompt) {
      setError("请输入提示词");
      return;
    }
    setGeneratedText("");
    setShowResult(false);
    setLoading(true);
    setError("");

    try {
      const response = await axios.post(
        DEEPSEEK_API_URL,
        {
          model: DEEPSEEK_MODEL,
          messages: [{ role: "user", content: trimmedPrompt }],
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

      const content = response.data?.choices?.[0]?.message?.content;
      if (!content) {
        throw new Error("API 返回了空响应");
      }

      setGeneratedText(content);
      // 延迟显示结果以触发动画
      setTimeout(() => setShowResult(true), 50);
    } catch (err: unknown) {
      handleApiError(err);
    } finally {
      setLoading(false);
    }
  };

  const onInsert = async () => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.insertText(generatedText, "Start");
        await context.sync();
      });
    } catch (err) {
      setError("插入文档失败，请确保 Word 文档已打开");
      console.error("Word insert error:", err);
    }
  };

  const onCopy = async () => {
    try {
      await navigator.clipboard.writeText(generatedText);
    } catch (err) {
      setError("复制到剪贴板失败");
      console.error("Clipboard copy error:", err);
    }
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
              width={36}
              height={36}
              decoding="async"
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
                <span className="char-count">
                  {prompt.length} / {MAX_PROMPT_LENGTH}
                </span>
              </div>
              <TextField
                placeholder="请输入..."
                value={prompt}
                rows={20}
                multiline={true}
                resizable={false}
                maxLength={MAX_PROMPT_LENGTH}
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
                iconProps={{ iconName: "Play" }}
                onClick={onClick}
                disabled={loading || !prompt.trim()}
                styles={{
                  root: {
                    background: "linear-gradient(135deg, #007aff 0%, #0a84ff 50%, #0060df 100%)",
                    color: "white",
                    margin: "20px 0 16px",
                    borderRadius: 22,
                    padding: "0 28px",
                    minWidth: 140,
                    height: 40,
                    border: "none",
                    position: "relative",
                    overflow: "hidden",
                    boxShadow: "0 4px 14px rgba(0, 122, 255, 0.35)",
                    transition: "all 0.25s cubic-bezier(0.4, 0, 0.2, 1)",
                    cursor: loading || !prompt.trim() ? "not-allowed" : "pointer",
                    opacity: loading || !prompt.trim() ? 0.6 : 1,
                    selectors: {
                      ":hover": {
                        background: "linear-gradient(135deg, #0066d9 0%, #007aff 50%, #0055b3 100%)",
                        boxShadow: "0 6px 20px rgba(0, 122, 255, 0.45)",
                        transform: loading || !prompt.trim() ? "none" : "translateY(-1px) scale(1.02)",
                      },
                      ":active": {
                        background: "linear-gradient(135deg, #0055b3 0%, #0060df 50%, #004499 100%)",
                        boxShadow: "0 2px 6px rgba(0, 122, 255, 0.3)",
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
                          "linear-gradient(135deg, transparent 0%, rgba(255,255,255,0.15) 50%, transparent 100%)",
                        pointerEvents: "none",
                      },
                    },
                  },
                  icon: {
                    color: "white",
                    fontSize: 14,
                    marginRight: 6,
                  },
                  label: {
                    fontWeight: 600,
                    fontSize: 14,
                    letterSpacing: "0.5px",
                  },
                  flexContainer: {
                    justifyContent: "center",
                    alignItems: "center",
                  },
                }}
              >
                {loading ? "生成中..." : "生成"}
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
                        background: "rgba(0, 122, 255, 0.08)",
                      },
                      icon: { color: "#007aff", fontSize: 14 },
                      label: { fontWeight: 600, fontSize: 13, color: "#007aff" },
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
                        background: "rgba(0, 122, 255, 0.08)",
                      },
                      icon: { color: "#007aff", fontSize: 14 },
                      label: { fontWeight: 600, fontSize: 13, color: "#007aff" },
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
