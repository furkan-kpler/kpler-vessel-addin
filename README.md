# Kpler Vessel Ownership & Particulars — Excel Add-in

A task-pane add-in for Excel (desktop and Excel on the web). The user pastes a list
of IMO numbers, the add-in queries the Kpler Vessel Ownership & Particulars GraphQL
API, and writes the full result set into a table on the active workbook.

## What's in this package

```
manifest.xml     Office Add-in manifest (classic XML manifest, works on Excel
                  desktop, Mac, and Excel on the web)
taskpane.html     Task pane UI
taskpane.css      Styling
taskpane.js       All logic: settings, GraphQL query/pagination, writing to the sheet
commands.html     Required "function file" for the ribbon button
commands.js
assets/           Placeholder ribbon icons (16/32/64/80px) — swap for your own artwork
proxy/            Cloudflare Worker CORS proxy - deploy this first, see proxy/README.md
```

## Connection settings

- **Proxy URL**: the Cloudflare Worker URL from `proxy/README.md` (e.g.
  `https://kpler-vessel-proxy.<subdomain>.workers.dev`) - **not** the raw Kpler
  endpoint. Kpler's endpoint doesn't return CORS headers browsers require, so a
  direct call from the task pane gets blocked with a generic "Load failed" even
  though the identical request works fine from `curl`. See `proxy/README.md` for
  the full diagnosis and the one-command deploy that fixes it.
- **Auth header**: pick the scheme that matches the customer's Kpler credential.
  - **Basic** (default) — paste the token exactly as given, either just the base64
    blob or the whole `Basic <token>` string; the add-in passes it through as-is,
    no re-encoding, no objection either way.
  - Bearer / raw Authorization / x-api-key are also available if a given customer's
    contract uses one of those instead.
- The key is stored in this browser's `localStorage` only — never written into the
  workbook, and never seen by the proxy worker's code/logs (it just relays the
  header through).

## Architecture

```
Excel task pane -> Cloudflare Worker proxy (proxy/worker.js) -> Kpler GraphQL API
```

The proxy exists solely to add CORS headers Kpler's endpoint doesn't send. It
carries no secret of its own; the customer's token still flows from the add-in on
every request. See `proxy/README.md` to deploy it (one `wrangler deploy` command).

## Fields returned

Every field exposed by the schema's non-deprecated types: identifiers (IMO, MMSI,
call sign, ENI, ship ID), full particulars (general, hull, dimensions, tonnage,
capacity, engine, fuel, AIS transceiver class), ownership (ultimate beneficial
owner(s), beneficial owner, registered owner), commercial (disponent owner,
commercial operator), management (technical manager, ISM manager), associated
companies (ship builder, engine builder, classification society, P&I club), and
historical vessel names.

## Deployment (Kubernetes + Kong)

The browser cannot call `api.kpler.com` directly (no CORS header on its preflight), so a small
pass-through proxy sits between the add-in and Kpler.

- `proxy-k8s/`: the containerized Node proxy (no npm dependencies). Build and test instructions are in [proxy-k8s/README.md](proxy-k8s/README.md).
- `deploy/solution-eng-gitops/`: Helm values in the layout of `Kpler/solution-eng-gitops`. Copy `gitops/vessel-ownership-proxy/` into that repo.
- `deploy/kong-gitops/`: Kong service config in the layout of `Kpler/kong-gitops`. Copy `kong-config/services/vessel-ownership-proxy/` into that repo.

Open decisions (tracked as `TODO(team)` in the files): exposure type, the production hostname,
the allowed add-in origins, `authorizedSources`, and the image registry and tag.

## Pilot setup (run locally, desktop Excel)

This runs the whole add-in on your own machine. No deployment is needed. Excel on the web can't load a `localhost` add-in, so this is for desktop Excel only.

Requirements: Node 22 or newer.

1. Install the Office dev certificates once (trust prompt, admin rights may be needed):
   ```bash
   npx office-addin-dev-certs install
   ```
2. Terminal 1, start the proxy on `http://localhost:8080`:
   ```bash
   npm run proxy
   ```
3. Terminal 2, serve the add-in on `https://localhost:3000`:
   ```bash
   npm run addin
   ```
4. In Excel, go to Insert → Add-ins → My Add-ins → Upload My Add-in, and choose `manifest-local.xml` from this folder.
5. In the task pane's Connection settings:
   - Proxy URL: `http://localhost:8080`
   - Auth header: `Basic`
   - API key / token: your Kpler API key

Use your own key only, never one that belongs to someone else. The proxy only forwards it to Kpler and keeps nothing.

To check the proxy, run `npm run test:proxy`.
