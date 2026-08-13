const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const MiniCssExtractPlugin = require("mini-css-extract-plugin");
const devCerts = require("office-addin-dev-certs");

const urlDev = "https://localhost:3000/";
const urlProd = "https://sunnyliu2025.github.io/WordGPT/";

module.exports = async (env, argv) => {
  const isDev = argv.mode === "development";
  const url = isDev ? urlDev : urlProd;

  // 开发模式下，让 dev-server 的端口与 HTTPS 和 publicPath(https://localhost:3000/) 保持一致，
  // 避免页面引用的 JS/CSS 指向 3000 而 dev-server 却运行在默认端口 8080 导致资源加载失败。
  let devServerConfig = {};
  if (isDev) {
    try {
      devServerConfig = {
        https: await devCerts.getHttpsServerOptions(),
        port: 3000,
        host: "localhost",
      };
    } catch (err) {
      console.warn(
        "无法加载 office-addin-dev-certs 证书，dev-server 将退回默认配置。",
        err
      );
    }
  }

  return {
    target: "web",
    devtool: "source-map",
    devServer: devServerConfig,
    entry: {
      polyfill: {
        import: ["core-js/stable", "regenerator-runtime/runtime"],
      },
      vendor: {
        import: ["react", "react-dom", "@fluentui/react", "axios"],
        dependOn: "polyfill",
      },
      taskpane: {
        import: "./src/taskpane/index.tsx",
        dependOn: "vendor",
      },
      commands: {
        import: "./src/commands/commands.ts",
        dependOn: "vendor",
      },
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      publicPath: url,
      filename: "[name].js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js", ".jsx", ".json"],
      fallback: {
        os: require.resolve("os-browserify/browser"),
        process: require.resolve("process/browser"),
      },
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: {
            loader: "ts-loader",
            options: {
              transpileOnly: true,
            },
          },
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: [MiniCssExtractPlugin.loader, "css-loader"],
        },
        {
          test: /\.less$/,
          use: [MiniCssExtractPlugin.loader, "css-loader", "less-loader"],
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico|svg)$/,
          type: "asset/resource",
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
      ],
    },
    plugins: [
      new MiniCssExtractPlugin({
        filename: "[name].[contenthash].css",
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "vendor", "taskpane"],
        scriptLoading: "defer",
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
        scriptLoading: "defer",
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets",
            to: "assets",
          },
          {
            from: "index.html",
            to: "index.html",
            noErrorOnMissing: true,
          },
        ],
      }),
    ],
    performance: {
      hints: false,
    },
  };
};
