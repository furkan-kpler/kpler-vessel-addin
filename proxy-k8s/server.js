/**
 * Kpler Vessel API CORS proxy (containerized, for Kubernetes).
 *
 * Why this exists: the Kpler GraphQL endpoint does not return
 * Access-Control-Allow-Origin on its preflight response, so the Excel task pane
 * (a browser context) is blocked before the real request goes out.
 *
 * What it does:
 * - Forwards POST requests (body + Authorization header) to ONE fixed upstream
 *   URL and returns the upstream response, adding CORS headers.
 * - Adds CORS only for origins listed in ALLOWED_ORIGINS. Other browser origins
 *   are refused with 403.
 *
 * What it does NOT do:
 * - It holds no credential. The caller's token travels in the Authorization
 *   header, exactly as the add-in sends it. Nothing is stored or logged.
 * - It never follows a caller-supplied URL, so it cannot be used to reach other
 *   hosts.
 *
 * Dependency-free: uses only Node built-ins (Node 22+, global fetch).
 */

"use strict";

const fs = require("node:fs");
const http = require("node:http");
const https = require("node:https");
const { URL } = require("node:url");

const DEFAULTS = {
  upstream: "https://api.kpler.com/v2/maritime/vessels/graphql",
  port: 8080,
  allowedOrigins: [],
  maxBodyBytes: 1_000_000,
  upstreamTimeoutMs: 60_000,
};

/** Read config from environment variables (see README for meanings). */
function configFromEnv(env = process.env) {
  const upstream = env.KPLER_ENDPOINT || DEFAULTS.upstream;
  const parsed = new URL(upstream); // throws on a malformed value, so bad config fails fast
  if (parsed.protocol !== "https:" && !(parsed.hostname === "localhost" || parsed.hostname === "127.0.0.1")) {
    throw new Error("KPLER_ENDPOINT must be https (plain http is only allowed for localhost tests).");
  }
  return {
    upstream,
    port: Number(env.PORT || DEFAULTS.port),
    // Unset binds all interfaces (right for a container). Set HOST=127.0.0.1 to listen on this machine only.
    host: env.HOST || undefined,
    allowedOrigins: (env.ALLOWED_ORIGINS || "")
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean),
    maxBodyBytes: Number(env.MAX_BODY_BYTES || DEFAULTS.maxBodyBytes),
    upstreamTimeoutMs: Number(env.UPSTREAM_TIMEOUT_MS || DEFAULTS.upstreamTimeoutMs),
    // Optional TLS, so an https add-in page can call a localhost proxy without mixed-content blocking.
    tls:
      env.TLS_KEY_FILE && env.TLS_CERT_FILE
        ? { key: fs.readFileSync(env.TLS_KEY_FILE), cert: fs.readFileSync(env.TLS_CERT_FILE) }
        : null,
  };
}

/**
 * Returns CORS headers for a request origin, or null if the origin is not allowed.
 * A request with no Origin header is not a browser request, so CORS does not apply.
 */
function corsHeadersFor(origin, allowedOrigins) {
  const base = {
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Access-Control-Max-Age": "86400",
    Vary: "Origin",
  };
  if (allowedOrigins.includes("*")) {
    return { ...base, "Access-Control-Allow-Origin": "*" };
  }
  if (origin && allowedOrigins.includes(origin)) {
    return { ...base, "Access-Control-Allow-Origin": origin };
  }
  return null;
}

function send(res, status, body, headers = {}) {
  const payload = body === undefined ? "" : typeof body === "string" ? body : JSON.stringify(body);
  res.writeHead(status, {
    "Content-Type": "application/json",
    "Content-Length": Buffer.byteLength(payload),
    ...headers,
  });
  res.end(payload);
}

function sendError(res, status, message, headers = {}) {
  send(res, status, { errors: [{ message }] }, headers);
}

/** Collect the request body, refusing anything over maxBytes. */
function readBody(req, maxBytes) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let total = 0;
    req.on("data", (chunk) => {
      total += chunk.length;
      if (total > maxBytes) {
        // Keep draining (without buffering) so the 413 response can still be sent.
        reject(Object.assign(new Error("Request body too large."), { status: 413 }));
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

function createProxyServer(config) {
  const cfg = { ...DEFAULTS, ...config };

  const handler = async (req, res) => {
    const started = Date.now();
    const origin = req.headers.origin;
    const path = new URL(req.url, "http://localhost").pathname;

    // Liveness/readiness probes used by Kubernetes. No auth, no CORS needed.
    if (req.method === "GET" && (path === "/healthz" || path === "/readyz")) {
      return send(res, 200, { status: "ok" });
    }

    const cors = corsHeadersFor(origin, cfg.allowedOrigins);

    // A browser call from an origin we do not allow is refused outright.
    if (origin && !cors) {
      log(req, 403, started);
      return sendError(res, 403, "Origin not allowed.");
    }

    if (req.method === "OPTIONS") {
      // Chromium asks before a public page (the GitHub Pages add-in) calls a local address.
      const preflightHeaders = { ...(cors || {}) };
      if (req.headers["access-control-request-private-network"] === "true") {
        preflightHeaders["Access-Control-Allow-Private-Network"] = "true";
      }
      res.writeHead(204, preflightHeaders);
      res.end();
      log(req, 204, started);
      return;
    }

    if (req.method !== "POST") {
      log(req, 405, started);
      return sendError(res, 405, "Method not allowed - POST only.", cors || {});
    }

    const auth = req.headers.authorization;
    if (!auth) {
      log(req, 401, started);
      return sendError(res, 401, "Missing Authorization header.", cors || {});
    }

    let body;
    try {
      body = await readBody(req, cfg.maxBodyBytes);
    } catch (err) {
      const status = err.status || 400;
      log(req, status, started);
      return sendError(res, status, `Could not read request body: ${err.message}`, cors || {});
    }

    let upstreamRes;
    try {
      upstreamRes = await fetch(cfg.upstream, {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: auth },
        body,
        signal: AbortSignal.timeout(cfg.upstreamTimeoutMs),
      });
    } catch (err) {
      log(req, 502, started);
      return sendError(res, 502, `Upstream request to Kpler failed: ${err.message}`, cors || {});
    }

    const text = await upstreamRes.text();
    const contentType = upstreamRes.headers.get("content-type") || "application/json";
    res.writeHead(upstreamRes.status, { "Content-Type": contentType, ...(cors || {}) });
    res.end(text);
    log(req, upstreamRes.status, started);
  };
  return cfg.tls ? https.createServer(cfg.tls, handler) : http.createServer(handler);
}

/** One structured log line per request. Never logs headers or bodies. */
function log(req, status, started) {
  const line = {
    time: new Date().toISOString(),
    method: req.method,
    path: req.url ? req.url.split("?")[0] : "",
    status,
    ms: Date.now() - started,
  };
  console.log(JSON.stringify(line));
}

module.exports = { configFromEnv, createProxyServer, corsHeadersFor };

// Start the server when run directly (node server.js / container CMD).
if (require.main === module) {
  const config = configFromEnv();
  const server = createProxyServer(config);
  const listenArgs = config.host ? [config.port, config.host] : [config.port];
  server.listen(...listenArgs, () => {
    console.log(
      JSON.stringify({
        msg: "listening",
        url: `${config.tls ? "https" : "http"}://localhost:${config.port}`,
        upstream: config.upstream,
      })
    );
  });
  const shutdown = () => server.close(() => process.exit(0));
  process.on("SIGTERM", shutdown);
  process.on("SIGINT", shutdown);
}
