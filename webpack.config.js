/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const MiniCssExtractPlugin = require("mini-css-extract-plugin");
const CssMinimizerPlugin = require("css-minimizer-webpack-plugin");
const TerserPlugin = require("terser-webpack-plugin");
const webpack = require("webpack");

const urlDev = "https://localhost:3000/";
const urlProd = "https://sunnyliu2025.github.io/WordGPT/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    // 生产环境不生成 source-map，开发环境使用轻量级 eval-source-map
    devtool: dev ? "eval-source-map" : false,
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      // 注意：不再手动声明 vendor entry，改由 splitChunks 自动提取第三方依赖
      // react-hot-loader/patch 仅在开发环境需要
      taskpane: [
        ...(dev ? ["react-hot-loader/patch"] : []),
        "./src/taskpane/index.tsx",
        "./src/taskpane/taskpane.html",
      ],
      commands: "./src/commands/commands.ts",
    },
    output: {
      clean: true,
      filename: dev ? "[name].js" : "[name].[contenthash:8].js",
      chunkFilename: dev ? "[name].js" : "[name].[contenthash:8].js",
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
      // 开发环境使用 react-hot-loader 的 react-dom 以支持 HMR
      ...(dev && {
        alias: {
          "react-dom": "@hot-loader/react-dom",
        },
      }),
    },
    optimization: {
      // 关键：提取所有第三方依赖到 vendor chunk，避免重复打包
      splitChunks: {
        chunks: "all",
        cacheGroups: {
          vendor: {
            name: "vendor",
            test: /[\\/]node_modules[\\/]/,
            chunks: "all",
            priority: 10,
          },
          // Fluent UI 体积较大，单独拆分为 fluentui chunk，可独立缓存
          fluentui: {
            name: "fluentui",
            test: /[\\/]node_modules[\\/]@fluentui[\\/]/,
            chunks: "all",
            priority: 20,
            reuseExistingChunk: true,
          },
        },
      },
      minimizer: [
        new TerserPlugin({
          terserOptions: {
            compress: {
              drop_console: !dev, // 生产环境移除 console.log
              drop_debugger: true,
            },
            output: {
              comments: false, // 移除注释
            },
          },
          extractComments: false, // 不提取 LICENSE 文件
        }),
        new CssMinimizerPlugin(), // 压缩 CSS
      ],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-typescript"],
            },
          },
        },
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: [...(dev ? ["react-hot-loader/webpack"] : []), "ts-loader"],
        },
        {
          test: /\.css$/,
          use: [
            MiniCssExtractPlugin.loader,
            {
              loader: "css-loader",
              options: {
                // 生产环境启用 CSS 压缩
                ...(!dev && {
                  importLoaders: 0,
                }),
              },
            },
          ],
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: [
            {
              loader: "html-loader",
              options: {
                // 生产环境压缩 HTML
                minimize: !dev,
                sources: false, // 不处理 HTML 中的资源引用（由 webpack 处理）
              },
            },
          ],
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
        {
          test: /\.(woff|woff2|eot|ttf|svg)$/,
          type: "asset/resource",
          generator: {
            filename: "fonts/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
      new MiniCssExtractPlugin({
        filename: dev ? "[name].css" : "[name].[contenthash:8].css",
        chunkFilename: dev ? "[name].css" : "[name].[contenthash:8].css",
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane", "vendor", "fluentui", "polyfill"],
        // 生产环境压缩 HTML
        ...(!dev && {
          minify: {
            removeComments: true,
            collapseWhitespace: true,
            removeAttributeQuotes: true,
          },
        }),
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
        ...(!dev && {
          minify: {
            removeComments: true,
            collapseWhitespace: true,
            removeAttributeQuotes: true,
          },
        }),
      }),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
    ],
    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
