#!/usr/bin/env node
"use strict";

// Serves the add-in's static files over HTTPS on localhost so Excel can load them when sideloaded.
// Needs Office dev certificates first: `npx office-addin-dev-certs install`.
// Only the add-in files are served, not the rest of the repo.

const fs = require("node:fs");
const https = require("node:https");
const os = require("node:os");
const path = require("node:path");

const ROOT = path.resolve(__dirname, "..");
const PORT = Number(process.env.ADDIN_PORT || 3000);
const CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");
const KEY_FILE = path.join(CERT_DIR, "localhost.key");
const CERT_FILE = path.join(CERT_DIR, "localhost.crt");

const SERVED_FILES = new Set(["taskpane.html", "taskpane.js", "taskpane.css", "commands.html", "commands.js"]);
const CONTENT_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
};

function resolveRequest(urlPath) {
  const clean = decodeURIComponent(urlPath.split("?")[0]);
  const rel = clean === "/" ? "taskpane.html" : clean.replace(/^\/+/, "");
  if (!SERVED_FILES.has(rel) && !rel.startsWith("assets/")) return null;
  const full = path.resolve(ROOT, rel);
  if (!full.startsWith(ROOT + path.sep)) return null;
  return full;
}

if (!fs.existsSync(KEY_FILE) || !fs.existsSync(CERT_FILE)) {
  console.error(`Office dev certificates not found in ${CERT_DIR}.`);
  console.error("Run once: npx office-addin-dev-certs install");
  process.exit(1);
}

const server = https.createServer(
  { key: fs.readFileSync(KEY_FILE), cert: fs.readFileSync(CERT_FILE) },
  (req, res) => {
    if (req.method !== "GET" && req.method !== "HEAD") {
      res.writeHead(405).end();
      return;
    }
    let file = null;
    try {
      file = resolveRequest(req.url || "/");
    } catch {
      file = null;
    }
    if (!file || !fs.existsSync(file) || !fs.statSync(file).isFile()) {
      res.writeHead(404).end("Not found");
      return;
    }
    res.writeHead(200, {
      "Content-Type": CONTENT_TYPES[path.extname(file)] || "application/octet-stream",
      "Cache-Control": "no-store",
    });
    if (req.method === "HEAD") {
      res.end();
      return;
    }
    fs.createReadStream(file).pipe(res);
  }
);

server.listen(PORT, "localhost", () => {
  console.log(`Add-in served at https://localhost:${PORT}/taskpane.html`);
});
