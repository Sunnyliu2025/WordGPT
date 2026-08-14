import * as React from "react";
import {
  CommandButton,
  DefaultButton,
  IconButton,
  MessageBar,
  MessageBarType,
} from "@fluentui/react";
import type { IButtonStyles } from "@fluentui/react";
import axios from "axios";
import Center from "./Center";
import Container from "./Container";
import Login from "./Login";
import "./initializeIcons";
/* global Word, localStorage, navigator, console, setInterval, clearInterval, setTimeout */

const MAX_PROMPT_LENGTH = 4000;
const DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions";
const DEEPSEEK_MODEL = "deepseek-v4-flash";
// 上传文件大小上限：10MB
const MAX_FILE_SIZE = 10 * 1024 * 1024;
// 附加文件内容的最大字符数（避免超出模型上下文限制）
const MAX_FILE_CONTENT_CHARS = 20000;

interface ErrorResponse {
  message?: string;
}

interface AttachedFile {
  name: string;
  size: number;
  content: string;
  truncated: boolean;
}

/* ===== 静态样式常量：避免每次渲染都重建对象，降低重渲染开销 ===== */

const CLEAR_BUTTON_STYLES: IButtonStyles = {
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
};

const TEXTAREA_STYLE: React.CSSProperties = {
  display: "block",
  width: "100%",
  minWidth: "100%",
  maxWidth: "100%",
  minHeight: "160px",
  maxHeight: "400px",
  height: "160px",
  boxSizing: "border-box",
  padding: "16px 18px",
  fontSize: "15px",
  fontFamily:
    "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
  lineHeight: "32px",
  color: "#1f2937",
  background: "#f9fafb",
  border: "2px solid #e5e7eb",
  borderRadius: "12px",
  outline: "none",
  resize: "none",
  overflowY: "auto",
  overflowX: "hidden",
  whiteSpace: "pre-wrap",
  wordBreak: "break-word",
  overflowWrap: "break-word",
};

const REMOVE_FILE_BUTTON_STYLES: IButtonStyles = {
  root: {
    color: "#9ca3af",
    width: 28,
    height: 28,
    borderRadius: 6,
    transition: "all 0.2s ease",
  },
  rootHovered: {
    color: "#dc2626",
    background: "rgba(220, 38, 38, 0.08)",
  },
  icon: { fontSize: 12, fontWeight: 700 },
};

const FILE_UPLOAD_BUTTON_STYLES: IButtonStyles = {
  root: {
    width: "100%",
    height: 40,
    borderRadius: 10,
    border: "1.5px dashed #cbd5e1",
    background: "#f8fafc",
    color: "#64748b",
    transition: "all 0.2s ease",
  },
  rootHovered: {
    background: "#f0f5ff",
    borderColor: "#3b82f6",
    color: "#3b82f6",
  },
  flexContainer: {
    justifyContent: "center",
    alignItems: "center",
  },
  label: { fontWeight: 500, fontSize: 13 },
};

const ACTION_BUTTON_STYLES: IButtonStyles = {
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
  label: {
    fontWeight: 600,
    fontSize: 13,
    color: "#007aff",
  },
};

/** 构建“生成”按钮样式（依赖 loading / promptReady，仅在二者变化时重建） */
const buildGenerateButtonStyles = (
  loading: boolean,
  promptReady: boolean
): IButtonStyles => ({
  root: {
    background:
      "linear-gradient(135deg, #007aff 0%, #0a84ff 50%, #0060df 100%)",
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
    cursor: loading || !promptReady ? "not-allowed" : "pointer",
    opacity: loading || !promptReady ? 0.6 : 1,
    selectors: {
      ":hover": {
        background:
          "linear-gradient(135deg, #0066d9 0%, #007aff 50%, #0055b3 100%)",
        boxShadow: "0 6px 20px rgba(0, 122, 255, 0.45)",
        transform:
          loading || !promptReady ? "none" : "translateY(-1px) scale(1.02)",
      },
      ":active": {
        background:
          "linear-gradient(135deg, #0055b3 0%, #0060df 50%, #004499 100%)",
        boxShadow: "0 2px 6px rgba(0, 122, 255, 0.3)",
        transform:
          loading || !promptReady ? "none" : "translateY(0) scale(0.98)",
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
});

export default function App() {
  const [apiKey, setApiKey] = React.useState<string>("");
  // 标记首次初始化是否完成（读取 localStorage 前不渲染主界面，避免“登录页一闪而过”的抖动）
  const [initialized, setInitialized] = React.useState<boolean>(false);
  const [prompt, setPrompt] = React.useState<string>("");
  const [error, setError] = React.useState<string>("");
  const [loading, setLoading] = React.useState<boolean>(false);
  const [generatedText, setGeneratedText] = React.useState<string>("");

  // 上传的附件文件
  const [attachedFile, setAttachedFile] = React.useState<AttachedFile | null>(
    null
  );
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  // 动画状态
  const [showResult, setShowResult] = React.useState<boolean>(false);
  const [dots, setDots] = React.useState<string>("");

  // 防重复提交 & 防止过期响应覆盖新结果
  const submitGuardRef = React.useRef<boolean>(false);
  const requestIdRef = React.useRef<number>(0);

  React.useEffect(() => {
    try {
      const key = localStorage.getItem("apiKey");
      if (key) {
        setApiKey(key);
      }
    } catch (err) {
      console.error("读取 localStorage 失败:", err);
    } finally {
      setInitialized(true);
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
      const axiosErr = error as {
        response?: { status?: number; data?: ErrorResponse };
        message?: string;
      };
      const status = axiosErr.response?.status;
      const data = axiosErr.response?.data;
      setError(
        `错误 ${status || "未知"}: ${
          data?.message || axiosErr.message || "未知错误"
        }`
      );
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

  /** 读取文件为文本（兼容旧版 WebView） */
  const readFileAsText = (file: File): Promise<string> =>
    new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || ""));
      reader.onerror = () => reject(reader.error);
      reader.readAsText(file, "utf-8");
    });

  /** 读取文件为 ArrayBuffer（兼容旧版 WebView） */
  const readFileAsArrayBuffer = (file: File): Promise<ArrayBuffer> =>
    new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(file);
    });

  /** 处理上传的文件，解析为纯文本内容 */
  const handleFileSelect = async (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = event.target.files?.[0];
    // 重置 input 值，允许重复选择同一个文件
    event.target.value = "";
    if (!file) return;

    if (file.size > MAX_FILE_SIZE) {
      setError("文件大小不能超过 10MB");
      return;
    }

    const lowerName = file.name.toLowerCase();
    try {
      let content = "";
      if (lowerName.endsWith(".docx")) {
        // 使用 mammoth 解析 Word 文档
        const mammoth = await import("mammoth");
        const arrayBuffer = await readFileAsArrayBuffer(file);
        const result = await mammoth.extractRawText({ arrayBuffer });
        content = result.value;
      } else if (lowerName.endsWith(".doc")) {
        setError("暂不支持旧版 .doc 格式，请转换为 .docx 后重试");
        return;
      } else if (lowerName.endsWith(".pdf")) {
        setError("暂不支持 PDF 文件，请转换为文本或 .docx 后重试");
        return;
      } else {
        // 其他文件按纯文本读取
        content = await readFileAsText(file);
      }

      const truncated = content.length > MAX_FILE_CONTENT_CHARS;
      const finalContent = truncated
        ? content.slice(0, MAX_FILE_CONTENT_CHARS)
        : content;

      setAttachedFile({
        name: file.name,
        size: file.size,
        content: finalContent,
        truncated,
      });
      setError("");
    } catch (err) {
      setError("文件读取失败，请重试");
      console.error("File read error:", err);
    }
  };

  /** 移除已上传的文件 */
  const onRemoveFile = () => {
    setAttachedFile(null);
  };

  /** 构建最终提交给 API 的提示词内容（提示词 + 附件内容） */
  const buildPromptContent = () => {
    const trimmed = prompt.trim();
    if (attachedFile && attachedFile.content) {
      const fileBlock = `\n\n===== 附件文件内容（${attachedFile.name}）=====\n${attachedFile.content}`;
      return `${trimmed}${fileBlock}`;
    }
    return trimmed;
  };

  const onClick = async () => {
    // 防止重复提交（按钮 disabled 之外的第二道保险）
    if (submitGuardRef.current) return;

    const trimmedPrompt = prompt.trim();
    const hasFileContent = !!attachedFile?.content;
    if (!trimmedPrompt && !hasFileContent) {
      setError("请输入提示词或上传文件");
      return;
    }

    submitGuardRef.current = true;
    const requestId = ++requestIdRef.current;
    setGeneratedText("");
    setShowResult(false);
    setLoading(true);
    setError("");

    try {
      const response = await axios.post(
        DEEPSEEK_API_URL,
        {
          model: DEEPSEEK_MODEL,
          messages: [{ role: "user", content: buildPromptContent() }],
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

      // 响应已过期（期间发起了新请求），直接丢弃，避免旧结果覆盖新结果
      if (requestId !== requestIdRef.current) return;

      const content = response.data?.choices?.[0]?.message?.content;
      if (!content) {
        throw new Error("API 返回了空响应");
      }

      setGeneratedText(content);
      // 延迟一帧设置可见状态，确保 opacity 过渡动画从 0 → 1 正常触发
      setTimeout(() => setShowResult(true), 50);
    } catch (err: unknown) {
      if (requestId !== requestIdRef.current) return;
      handleApiError(err);
    } finally {
      // 仅当是最新请求时才释放 loading 与提交锁
      if (requestId === requestIdRef.current) {
        setLoading(false);
        submitGuardRef.current = false;
      }
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
    setAttachedFile(null);
    setGeneratedText("");
    setShowResult(false);
    setError("");
  };

  // 是否有有效内容（提示词或附件），用于控制生成按钮状态
  const promptReady = !!prompt.trim() || !!attachedFile?.content;

  // 生成按钮样式仅在 loading / promptReady 变化时重建
  const generateButtonStyles = React.useMemo(
    () => buildGenerateButtonStyles(loading, promptReady),
    [loading, promptReady]
  );

  // 首次初始化完成前展示骨架屏，避免“登录页一闪而过”的抖动
  if (!initialized) {
    return (
      <Container>
        <div className="app-wrapper">
          <header className="app-header">
            <div className="app-header-content">
              <div className="skeleton skeleton-logo" />
              <div className="app-header-text">
                <div className="skeleton skeleton-title" />
                <div className="skeleton skeleton-subtitle" />
              </div>
            </div>
          </header>
          <div className="skeleton skeleton-block" />
          <div className="skeleton skeleton-block short" />
          <div className="skeleton skeleton-button" />
        </div>
      </Container>
    );
  }

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
            styles={CLEAR_BUTTON_STYLES}
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
              <textarea
                placeholder="输入你的提示词..."
                value={prompt}
                maxLength={MAX_PROMPT_LENGTH}
                onChange={(e) => setPrompt(e.target.value)}
                style={TEXTAREA_STYLE}
              />

              {/* ===== 文件上传区 ===== */}
              <div className="file-upload-section">
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".txt,.md,.csv,.json,.js,.jsx,.ts,.tsx,.py,.docx,.doc,text/plain"
                  style={{ display: "none" }}
                  onChange={handleFileSelect}
                />
                {attachedFile ? (
                  <div className="file-chip">
                    <span className="file-chip-icon">📄</span>
                    <div className="file-chip-info">
                      <span className="file-chip-name">
                        {attachedFile.name}
                      </span>
                      <span className="file-chip-meta">
                        {attachedFile.size >= 1024
                          ? `${(attachedFile.size / 1024).toFixed(1)} KB`
                          : `${attachedFile.size} B`}
                        {attachedFile.truncated ? " · 内容已截断" : ""}
                      </span>
                    </div>
                    <IconButton
                      iconProps={{ iconName: "Cancel" }}
                      title="移除文件"
                      ariaLabel="移除文件"
                      onClick={onRemoveFile}
                      styles={REMOVE_FILE_BUTTON_STYLES}
                    />
                  </div>
                ) : (
                  <DefaultButton
                    className="file-upload-btn"
                    onClick={() => fileInputRef.current?.click()}
                    styles={FILE_UPLOAD_BUTTON_STYLES}
                  >
                    <span className="file-upload-btn-content">
                      <span className="file-upload-btn-icon" aria-hidden="true">
                        📎
                      </span>
                      <span>上传文件（.txt / .docx / .md 等）</span>
                    </span>
                  </DefaultButton>
                )}
              </div>
            </div>

            {/* ===== 生成按钮 ===== */}
            <Center>
              <DefaultButton
                iconProps={{ iconName: "Play" }}
                onClick={onClick}
                disabled={loading || !promptReady}
                styles={generateButtonStyles}
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
              <div
                className={`result-section ${
                  showResult ? "result-visible" : ""
                }`}
              >
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
                    styles={ACTION_BUTTON_STYLES}
                  >
                    插入文档
                  </CommandButton>
                  <CommandButton
                    className="btn-action"
                    iconProps={{ iconName: "Copy" }}
                    onClick={onCopy}
                    styles={ACTION_BUTTON_STYLES}
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
