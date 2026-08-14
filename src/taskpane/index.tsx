import "./taskpane.css";
import App from "./components/App";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";
/* global Office, module, require, process */

// 生产环境使用 React.Fragment 替代 react-hot-loader 的 AppContainer：
// 避免在线上产物中引入运行时开销（dead-code elimination 会移除 require 分支）；
// 开发环境仍保留热更新能力。
const AppContainer: React.ComponentType<{ children?: React.ReactNode }> =
  process.env.NODE_ENV !== "production"
    ? (require("react-hot-loader") as any).AppContainer
    : (React.Fragment as React.ComponentType<{ children?: React.ReactNode }>);

const render = (Component: React.ComponentType) => {
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
  render(App);
});

// 开发环境热更新支持，生产环境自动移除
if (process.env.NODE_ENV !== "production" && (module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
