const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const MiniCssExtractPlugin = require("mini-css-extract-plugin");

const isProd = process.env.NODE_ENV === "production";

module.exports = {
  mode: isProd ? "production" : "development",
  devtool: isProd ? "source-map" : "inline-source-map",

  entry: {
    polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
    vendor: ["react", "react-dom", "@fluentui/react", "axios"],
    taskpane: path.resolve(__dirname, "src/taskpane/index.tsx"),
    commands: path.resolve(__dirname, "src/commands/commands.ts"),
  },

  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js",
    clean: true,
  },

  resolve: {
    extensions: [".ts", ".tsx", ".js", ".jsx"],
  },

  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: ["babel-loader", "ts-loader"],
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
        test: /\.html$/,
        use: ["html-loader"],
        exclude: /node_modules/,
      },
      {
        test: /\.(png|jpg|jpeg|gif|ico|svg)$/,
        type: "asset/resource",
        generator: {
          filename: "assets/[name][ext][query]",
        },
      },
    ],
  },

  plugins: [
    new MiniCssExtractPlugin({
      filename: "[name].[contenthash].css",
    }),
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: path.resolve(__dirname, "src/taskpane/taskpane.html"),
      chunks: ["polyfill", "vendor", "taskpane"],
    }),
    new HtmlWebpackPlugin({
      filename: "commands.html",
      template: path.resolve(__dirname, "src/commands/commands.html"),
      chunks: ["polyfill", "vendor", "commands"],
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "assets",
          to: "assets",
          noErrorOnMissing: true,
        },
        {
          from: "index.html",
          to: "index.html",
          noErrorOnMissing: true,
        },
      ],
    }),
  ],

  optimization: {
    splitChunks: {
      chunks: "all",
      cacheGroups: {
        vendor: {
          test: /[\\/]node_modules[\\/]/,
          name: "vendor",
          chunks: "all",
          priority: 10,
        },
      },
    },
  },

  devServer: {
    static: {
      directory: path.resolve(__dirname, "dist"),
    },
    hot: true,
    port: 3000,
    devMiddleware: {
      writeToDisk: true,
    },
  },
};
