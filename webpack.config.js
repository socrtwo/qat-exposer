const path = require("path");
const fs = require("fs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

// Use the certs installed by office-addin-dev-certs if available
function getHttpsConfig() {
  const home = process.env.HOME || process.env.USERPROFILE || "";
  const certDir = path.join(home, ".office-addin-dev-certs");
  const key = path.join(certDir, "localhost.key");
  const cert = path.join(certDir, "localhost.crt");
  const ca = path.join(certDir, "ca.crt");

  if (fs.existsSync(key) && fs.existsSync(cert)) {
    const options = { key: fs.readFileSync(key), cert: fs.readFileSync(cert) };
    if (fs.existsSync(ca)) options.ca = fs.readFileSync(ca);
    return { type: "https", options: options };
  }
  return "https";
}

module.exports = {
  entry: {
    taskpane: "./src/taskpane/taskpane.js",
    commands: "./src/commands/commands.js",
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js",
    clean: true,
  },
  devServer: {
    static: {
      directory: path.join(__dirname, "dist"),
    },
    port: 3000,
    server: getHttpsConfig(),
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: "./src/taskpane/taskpane.html",
      chunks: ["taskpane"],
    }),
    new HtmlWebpackPlugin({
      filename: "commands.html",
      template: "./src/commands/commands.html",
      chunks: ["commands"],
    }),
    new CopyWebpackPlugin({
      patterns: [
        { from: "src/assets", to: "assets" },
        { from: "manifests", to: "manifests" },
      ],
    }),
  ],
  resolve: {
    extensions: [".js"],
  },
};
