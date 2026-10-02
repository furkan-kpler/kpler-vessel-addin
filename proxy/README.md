# Kpler Vessel API CORS Proxy

Why this exists: the Kpler GraphQL endpoint (`api.kpler.com/v2/maritime/vessels/graphql`)
doesn't return `Access-Control-Allow-Origin` on its CORS preflight. Confirmed via:

```bash
curl -s -D - -o /dev/null -X OPTIONS "https://api.kpler.com/v2/maritime/vessels/graphql" \
  -H "Origin: https://furkan-kpler.github.io" \
  -H "Access-Control-Request-Method: POST" \
  -H "Access-Control-Request-Headers: content-type,authorization"
```

The response has `access-control-allow-headers` and `access-control-allow-methods`,
but no `access-control-allow-origin`. Browsers enforce that; `curl` doesn't - which
is why the identical request works from a terminal and fails with a generic
"Load failed" from inside Excel's task pane.

This worker sits in between, purely to fix that. It is a dumb pass-through:

```
Excel task pane  --POST, same body, same Authorization header-->  this worker
                                                                        |
                                                                        v
                                                        Kpler GraphQL endpoint
                                                        (server-to-server, no CORS)
                                                                        |
                                                                        v
Excel task pane  <--same response body, now with CORS headers added--  this worker
```

It holds **no credential of its own** - the customer's token still travels from the
add-in's Connection settings exactly as before, this worker just relays it and adds
the missing CORS headers to the response on the way back.

## Deploy (Cloudflare Workers - free tier covers this easily)

```bash
cd proxy
npm install -g wrangler      # one-time, if you don't have it
wrangler login                # opens a browser to authorize against your Cloudflare account
wrangler deploy
```

That prints a URL like:
```
https://kpler-vessel-proxy.<your-subdomain>.workers.dev
```

That's the URL to put in the add-in's **Proxy URL** field (Connection settings) -
not the raw Kpler endpoint anymore.

## Verify it directly before testing through Excel

```bash
curl -s -X POST "https://kpler-vessel-proxy.<your-subdomain>.workers.dev" \
  -H "Content-Type: application/json" \
  -H "Authorization: Basic <your token>" \
  -d @../test_body.json
```

Should return the same vessel JSON you got hitting Kpler directly. If it does, the
proxy is confirmed working and the only remaining variable is the add-in's own
request logic - not CORS, not auth shape.

## If you change the upstream Kpler URL later

Edit `KPLER_ENDPOINT` in `wrangler.toml`, then `wrangler deploy` again. No code
change needed.

## Cost / limits

Cloudflare Workers free tier: 100,000 requests/day, no credit card required. Each
vessel lookup batch is one request to this worker (which then may issue its own
paginated calls upstream only if a single IMO batch exceeds 1000 results) - this
should sit comfortably inside the free tier for normal customer usage.
