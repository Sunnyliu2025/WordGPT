const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const MiniCssExtractPlugin = require("mini-css-extract-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://sunnyliu2025.github.io/WordGPT/";

module.exports = (env, argv) => {
  const isDev = argv.mode === "development";
  const url = isDev ? urlDev : urlProd;

  return {
    target: "web",
    devtool: "source-map",
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
      path: path.resolve(__dirname, "."),
      publicPath: url,
      filename: "[name].js",
      clean: false,
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
