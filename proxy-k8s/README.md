# vessel-ownership-proxy (container)

A CORS pass-through for the Kpler vessel ownership GraphQL API. The add-in's task pane is a browser,
so it needs `Access-Control-Allow-Origin` on responses. This proxy adds it. It holds no credentials:
the caller's token is forwarded unchanged in the `Authorization` header.

## Behaviour

| Request | Result |
| --- | --- |
| `GET /healthz`, `GET /readyz` | 200 (used by Kubernetes probes) |
| `OPTIONS` from an allowed origin | 204 with CORS headers |
| `POST` from an origin not in `ALLOWED_ORIGINS` | 403 |
| `POST` with no `Authorization` | 401 |
| `POST` with a body over `MAX_BODY_BYTES` | 413 |
| `POST` otherwise | forwarded to `KPLER_ENDPOINT`, response returned with CORS headers |

The upstream is fixed by configuration. Callers cannot choose the destination, so the proxy cannot be used to reach other hosts.

## Configuration (environment variables)

| Variable | Default | Meaning |
| --- | --- | --- |
| `KPLER_ENDPOINT` | `https://api.kpler.com/v2/maritime/vessels/graphql` | Upstream URL. Must be https. |
| `ALLOWED_ORIGINS` | empty | Comma-separated browser origins allowed to call the proxy. `*` allows all. |
| `PORT` | `8080` | Listen port. |
| `MAX_BODY_BYTES` | `1000000` | Largest accepted request body. |
| `UPSTREAM_TIMEOUT_MS` | `60000` | Timeout for the upstream call. |

## Run and test (Node 22+)

```bash
npm test                      # runs test/server.test.js against a fake upstream
npm start                     # listens on PORT (default 8080)
```

## Build the image

```bash
docker build -t vessel-ownership-proxy:dev .
docker run --rm -p 8080:8080 -e ALLOWED_ORIGINS=https://furkan-kpler.github.io vessel-ownership-proxy:dev
curl -s localhost:8080/healthz
```

Before production: pin the base image by digest in `Dockerfile`, and set the registry path and tag to the ones the platform team confirms.

## Not done yet

- No rate limiting in the proxy. Kong or the platform should provide it before the service is exposed publicly.
- Not yet run in a cluster.
