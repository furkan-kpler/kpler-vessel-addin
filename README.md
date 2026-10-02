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
