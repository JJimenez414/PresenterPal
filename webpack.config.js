const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = {
  entry: {
    taskpane: "./src/taskpane.js",
    chatHistory: "./src/chatHistory.js"
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].bundle.js", // Changed from bundle.js to [name].bundle.js
    clean: true
  },
  devServer: {
    static: [{
      directory: path.join(__dirname, "dist"),
    }, {
      directory: path.join(__dirname, "src"),
      publicPath: "/"
    }],
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
    server: {
      type: 'https',
      options: {
        cert: process.env.SSL_CRT_FILE,
        key: process.env.SSL_KEY_FILE,
      }
    },
    port: 3000,
    host: 'localhost',
    allowedHosts: 'all',
    hot: true,
    open: false,
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: path.resolve(__dirname, "src/taskpane.html"),
      filename: "taskpane.html",
      chunks: ["chatHistory", "taskpane"], // Order matters - chatHistory needs to load first
      chunksSortMode: "manual"
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "assets",
          to: "assets"
        },
        {
          from: "manifest.xml",
          to: "manifest.xml"
        },
        {
          from: "src/*.css",
          to: "[name][ext]"
        }
      ]
    }),
  ],
  module: {
    rules: [
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
      },
    ],
  },
};