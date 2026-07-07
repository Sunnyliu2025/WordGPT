import * as React from "react";
import { DefaultButton, TextField } from "@fluentui/react";
import Center from "./Center";

interface LoginProps {
  onSave: (token: string) => void;
}

export default function Login({ onSave }: LoginProps) {
  const [token, setToken] = React.useState<string>("");

  const handleSave = () => {
    if (token.trim()) {
      onSave(token.trim());
    }
  };

  return (
    <div className="login-wrapper">
      {/* 装饰图标 */}
      <div className="login-icon-container">
        <div className="login-icon-ring">
          <span className="login-icon-emoji">🔑</span>
        </div>
      </div>

      {/* 标题 */}
      <div className="login-header">
        <h2 className="login-title">欢迎使用 WordGPT</h2>
        <p className="login-desc">请输入你的 DeepSeek API 密钥以开始使用</p>
      </div>

      {/* 输入区域 */}
      <div className="login-input-wrapper">
        <TextField
          value={token}
          onChange={(_, newValue: string) => setToken(newValue || "")}
          placeholder="sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
          styles={{
            root: { width: "100%" },
            fieldGroup: {
              borderRadius: 12,
              border: "2px solid #e5e7eb",
              transition: "border-color 0.2s ease, box-shadow 0.2s ease",
              background: "#f9fafb",
              height: 44,
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
              color: "#1f2937",
              padding: "0 14px",
              "::placeholder": {
                color: "#9ca3af",
                fontSize: 13,
              },
            },
          }}
        />
      </div>

      {/* 提示信息 */}
      <p className="login-hint">
        没有 API 密钥？
        <a
          href="https://platform.deepseek.com/api_keys"
          target="_blank"
          rel="noopener noreferrer"
          className="login-link"
        >
          前往 DeepSeek 获取
        </a>
      </p>

      {/* 按钮 */}
      <Center>
        <DefaultButton
          iconProps={{ iconName: "SaveAs" }}
          onClick={handleSave}
          disabled={!token.trim()}
          styles={{
            root: {
              minWidth: 180,
              height: 44,
              borderRadius: 12,
              border: "none",
              background: !token.trim()
                ? "#e5e7eb"
                : "linear-gradient(135deg, #2563eb 0%, #3b82f6 50%, #2563eb 100%)",
              color: !token.trim() ? "#9ca3af" : "#fff",
              fontWeight: 600,
              fontSize: 15,
              boxShadow: !token.trim()
                ? "none"
                : "0 4px 14px rgba(37, 99, 235, 0.35)",
              transition: "all 0.3s cubic-bezier(0.4, 0, 0.2, 1)",
              cursor: !token.trim() ? "not-allowed" : "pointer",
              selectors: {
                ":hover": {
                  background: !token.trim()
                    ? "#e5e7eb"
                    : "linear-gradient(135deg, #1d4ed8 0%, #2563eb 50%, #1d4ed8 100%)",
                  boxShadow: !token.trim()
                    ? "none"
                    : "0 6px 20px rgba(37, 99, 235, 0.45)",
                  transform: !token.trim() ? "none" : "translateY(-1px)",
                },
                ":active": {
                  transform: !token.trim() ? "none" : "translateY(0)",
                },
              },
            },
            icon: {
              color: !token.trim() ? "#9ca3af" : "#fff",
              fontSize: 14,
              marginRight: 8,
            },
            label: {
              letterSpacing: "0.5px",
            },
            flexContainer: {
              justifyContent: "center",
              alignItems: "center",
            },
          }}
        >
          保存 API 密钥
        </DefaultButton>
      </Center>
    </div>
  );
}
