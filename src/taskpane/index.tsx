import "./taskpane.css";
import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";
/* global Office, module, require, process */

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
