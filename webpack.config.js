const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");

module.exports = {
  entry: "./src/index.tsx",
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].[contenthash:8].js",
    clean: true,
  },
  resolve: {
    extensions: [".ts", ".tsx", ".js", ".json"],
    alias: {
      // 浏览器自包含版 exceljs，避免 webpack5 缺 node polyfill 报错
      exceljs: path.resolve(__dirname, "node_modules/exceljs/dist/exceljs.min.js"),
    },
  },
  module: {
    rules: [
      { test: /\.tsx?$/, use: "ts-loader", exclude: /node_modules/ },
      { test: /\.xlsx$/, type: "asset/resource" },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({ template: "./src/index.html" }),
  ],
  devServer: {
    port: 5173,
    headers: { "Access-Control-Allow-Origin": "*" },
    hot: true,
  },
  performance: { hints: false },
};
