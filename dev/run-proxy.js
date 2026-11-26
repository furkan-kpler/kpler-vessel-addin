#!/usr/bin/env node
"use strict";

// Starts the CORS proxy locally for the pilot, allowing the local add-in origin.

const { spawn } = require("node:child_process");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

// Use the Office dev certificates (run `npx office-addin-dev-certs install` once) so the proxy
// is served over https://localhost, which an https add-in page can call without mixed-content blocking.
const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
const keyFile = path.join(certDir, "localhost.key");
const certFile = path.join(certDir, "localhost.crt");
const useTls = fs.existsSync(keyFile) && fs.existsSync(certFile) && process.env.PROXY_HTTP !== "1";

const env = {
  ...process.env,
  PORT: process.env.PORT || "8080",
  // Local pilot: listen on this machine only, so nothing else on the network can reach the proxy.
  HOST: process.env.HOST || "127.0.0.1",
  ALLOWED_ORIGINS: process.env.ALLOWED_ORIGINS || "https://furkan-kpler.github.io,https://localhost:3000",
  ...(useTls ? { TLS_KEY_FILE: keyFile, TLS_CERT_FILE: certFile } : {}),
};

const child = spawn(process.execPath, [path.join(__dirname, "..", "proxy-k8s", "server.js")], {
  stdio: "inherit",
  env,
});
child.on("exit", (code) => process.exit(code ?? 0));
for (const signal of ["SIGINT", "SIGTERM"]) {
  process.on(signal, () => child.kill(signal));
}
