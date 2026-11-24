"use strict";

const { test, before, after } = require("node:test");
const assert = require("node:assert/strict");
const http = require("node:http");
const { createProxyServer, configFromEnv } = require("../server.js");

const ALLOWED = "https://addin.kpler.example";
let upstream, upstreamUrl, proxy, proxyUrl;
let lastUpstreamRequest = null;

before(async () => {
  // Fake Kpler endpoint: records what it received and answers with JSON.
  upstream = http.createServer((req, res) => {
    let body = "";
    req.on("data", (c) => (body += c));
    req.on("end", () => {
      lastUpstreamRequest = { headers: req.headers, body, method: req.method };
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ data: { ok: true } }));
    });
  });
  await new Promise((r) => upstream.listen(0, "127.0.0.1", r));
  upstreamUrl = `http://127.0.0.1:${upstream.address().port}/graphql`;

  proxy = createProxyServer({
    upstream: upstreamUrl,
    allowedOrigins: [ALLOWED],
    maxBodyBytes: 1000,
    upstreamTimeoutMs: 5000,
  });
  await new Promise((r) => proxy.listen(0, "127.0.0.1", r));
  proxyUrl = `http://127.0.0.1:${proxy.address().port}`;
});

after(() => {
  proxy.close();
  upstream.close();
});

test("health endpoints answer without auth", async () => {
  for (const path of ["/healthz", "/readyz"]) {
    const res = await fetch(proxyUrl + path);
    assert.equal(res.status, 200);
  }
});

test("preflight from an allowed origin returns CORS headers", async () => {
  const res = await fetch(proxyUrl + "/", {
    method: "OPTIONS",
    headers: { Origin: ALLOWED },
  });
  assert.equal(res.status, 204);
  assert.equal(res.headers.get("access-control-allow-origin"), ALLOWED);
  assert.match(res.headers.get("access-control-allow-headers"), /Authorization/);
});

test("request from a disallowed origin is refused with 403", async () => {
  const res = await fetch(proxyUrl + "/", {
    method: "POST",
    headers: { Origin: "https://evil.example", Authorization: "Basic abc", "Content-Type": "application/json" },
    body: "{}",
  });
  assert.equal(res.status, 403);
  assert.equal(res.headers.get("access-control-allow-origin"), null);
});

test("missing Authorization is rejected with 401 and never reaches Kpler", async () => {
  lastUpstreamRequest = null;
  const res = await fetch(proxyUrl + "/", {
    method: "POST",
    headers: { Origin: ALLOWED, "Content-Type": "application/json" },
    body: "{}",
  });
  assert.equal(res.status, 401);
  assert.equal(lastUpstreamRequest, null);
});

test("GET is not allowed", async () => {
  const res = await fetch(proxyUrl + "/", { method: "GET", headers: { Origin: ALLOWED } });
  assert.equal(res.status, 405);
});

test("valid POST is forwarded with the same body and Authorization, and CORS is added", async () => {
  lastUpstreamRequest = null;
  const payload = JSON.stringify({ query: "{ vessel { imo } }" });
  const res = await fetch(proxyUrl + "/", {
    method: "POST",
    headers: { Origin: ALLOWED, Authorization: "Basic dG9rZW4=", "Content-Type": "application/json" },
    body: payload,
  });
  assert.equal(res.status, 200);
  assert.equal(res.headers.get("access-control-allow-origin"), ALLOWED);
  assert.deepEqual(await res.json(), { data: { ok: true } });
  assert.equal(lastUpstreamRequest.method, "POST");
  assert.equal(lastUpstreamRequest.headers.authorization, "Basic dG9rZW4=");
  assert.equal(lastUpstreamRequest.body, payload);
});

test("oversized body is refused with 413", async () => {
  const res = await fetch(proxyUrl + "/", {
    method: "POST",
    headers: { Origin: ALLOWED, Authorization: "Basic abc", "Content-Type": "application/json" },
    body: "x".repeat(5000),
  });
  assert.equal(res.status, 413);
});

test("configFromEnv rejects a non-https upstream", () => {
  assert.throws(() => configFromEnv({ KPLER_ENDPOINT: "http://example.com/graphql" }), /must be https/);
  assert.doesNotThrow(() => configFromEnv({ KPLER_ENDPOINT: "https://api.kpler.com/v2/maritime/vessels/graphql" }));
});
