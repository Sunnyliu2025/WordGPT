import "./taskpane.css";
import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";
/* global Office, module, require, document, console */

const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <ThemeProvider>
        <Component />
      </ThemeProvider>
    </AppContainer>,
    // eslint-disable-next-line no-undef
    document.getElementById("container")
  );
};

Office.onReady(() => {
  try {
    render(App);
  } catch (e) {
    console.error("WordGPT render failed:", e);
    const container = document.getElementById("container");
    if (container) {
      container.innerHTML =
        '<div style="padding: 24px; font-family: -apple-system, sans-serif; color: #dc2626;">' +
        "<h2 style='font-size: 16px; margin: 0 0 8px;'>WordGPT 加载失败</h2>" +
        "<p style='font-size: 13px; margin: 0; color: #6b7280;'>请打开开发者工具查看控制台错误信息，然后重新加载加载项。</p>" +
        "</div>";
    }
  }
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
