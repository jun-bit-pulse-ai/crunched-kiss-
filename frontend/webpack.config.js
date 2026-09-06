/* eslint-disable no-undef */
const fs = require("fs");
const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://localhost:3000/";

function mkcertHttpsOptions() {
  const certDir = path.resolve(__dirname, "../certs");
  const candidates = [
    { key: "localhost-key.pem", cert: "localhost.pem" },
    { key: "localhost+2-key.pem", cert: "localhost+2.pem" },
  ];
  for (const files of candidates) {
    const key = path.join(certDir, files.key);
    const cert = path.join(certDir, files.cert);
    if (fs.existsSync(key) && fs.existsSync(cert)) {
      return { key: fs.readFileSync(key), cert: fs.readFileSync(cert) };
    }
  }
  return null;
}

async function getHttpsOptions() {
  const mkcert = mkcertHttpsOptions();
  if (mkcert) {
    return mkcert;
  }
  const devCerts = require("office-addin-dev-certs");
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (_env, options) => {
  const dev = options.mode === "development";
  return {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: "./src/taskpane/index.tsx",
      commands: "./src/commands/commands.ts",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: "ts-loader",
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg)$/,
          type: "asset/resource",
          generator: { filename: "assets/[name][ext]" },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets/*", to: "assets/[name][ext]" },
          {
            from: "manifest.xml",
            to: "[name][ext]",
            transform(content) {
              return dev ? content : content.toString().replaceAll(urlDev, urlProd);
            },
          },
        ],
      }),
    ],
    devServer: {
      hot: true,
      headers: { "Access-Control-Allow-Origin": "*" },
      server: {
        type: "https",
        options: await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };
};
