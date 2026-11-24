# vessel-ownership-proxy (Kong)

Exposes the `vessel-ownership-proxy` Kubernetes service (deployed from `solution-eng-gitops`)
through Kong at `/vessel-ownership-proxy`.

Status:
- dev (ireland): drafted, `internal-api.dev.galil.io`.
- prod: NOT drafted. The public host is a team decision (see the request to
  `#crew-data-ecosystem`). `api.kpler.com` is already used by the `vessel-api` route, so it cannot be reused.
