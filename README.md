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
```

## 1. Fill in the two things I couldn't know for you

I built this against the GraphQL schema for `vessel-ownership-and-particulars`
(confirmed via the schema introspection), but I don't have the customer's actual
endpoint URL or their auth header format — `developers.kpler.com` renders its spec
client-side so I couldn't scrape it, and that's exactly the kind of account-specific
detail (base URL, auth scheme, rate limit) the customer's onboarding docs / contract
will spell out precisely. Two things to confirm before handing this off:

1. **The GraphQL endpoint URL.** The customer enters this themselves in the task
   pane's "Connection settings" panel (see below) — no code change needed. If you'd
   rather it be pre-filled, set a default in `taskpane.js` at the top of
   `loadSettings()`.
2. **The auth header shape.** The task pane defaults to **Basic** auth
   (`Authorization: Basic base64(username:key)`, with a Username field that only
   shows when Basic is selected) and also offers `Bearer <key>`, raw
   `Authorization: <key>`, and `x-api-key: <key>`. Confirm which one the customer's
   Kpler contract uses and tell them which to pick (or hardcode it in
   `buildAuthHeaders()` in `taskpane.js` if it's always the same for every customer).

Also worth a quick check: if the endpoint enforces CORS restrictions on browser-origin
requests, calls from the task pane (which runs in a browser control inside Excel) may
need the add-in's hosting domain allow-listed on the API side — check with API/infra
if the customer hits a CORS error.

## 2. Host the files

Excel loads the add-in from a URL — it can't sideload local files directly. Put this
whole folder on any static HTTPS host (an existing Kpler-managed static site, GitHub
Pages, Azure Static Web Apps, S3+CloudFront, etc.), then replace every occurrence of
`REPLACE-WITH-YOUR-HOSTING-DOMAIN` in `manifest.xml` with the real domain
(4 places: `IconUrl`, `HighResolutionIconUrl`, `AppDomains`, `SourceLocation`, and the
two `bt:Url`/`bt:Image` entries under `Resources`).

## 3. Sideload it for the customer

- **Excel desktop**: Insert tab → Add-ins → Upload My Add-in → select `manifest.xml`.
- **Excel on the web**: Insert → Add-ins → Upload My Add-in → select `manifest.xml`.
- **Org-wide rollout**: an admin can deploy it centrally via the Microsoft 365 admin
  center (Integrated Apps) instead of per-user sideloading.

## 4. Using it

1. Open the add-in from the ribbon (Home tab → "Vessel Lookup").
2. Expand **Connection settings**, paste the endpoint URL and API key, pick the auth
   header format, click **Save settings**. This is stored in the browser's
   `localStorage` only — never written into the workbook, so sharing the file doesn't
   leak the key.
3. Paste IMO numbers into the box (one per line, or comma/space separated).
4. Optionally rename the output sheet (defaults to "Vessel Data").
5. Click **Fetch & write to sheet**. The add-in batches requests (250 IMOs per call,
   paginating automatically if a batch is large), then writes one row per vessel as
   an Excel Table with autofit columns, and reports any IMOs that had no match.

## Fields returned

Every field exposed by the schema's non-deprecated types: identifiers (IMO, MMSI,
call sign, ENI, ship ID), full particulars (general, hull, dimensions, tonnage,
capacity, engine, fuel, AIS transceiver class), ownership (ultimate beneficial
owner(s), beneficial owner, registered owner), commercial (disponent owner,
commercial operator), management (technical manager, ISM manager), associated
companies (ship builder, engine builder, classification society, P&I club), and
historical vessel names. Current company relationships are shown as
`Name / Country`; multiple UBOs are semicolon-joined with their share % and
confidence level. If the customer wants historical ownership (not just current) as
separate columns too, that's a small extension to the `COLUMNS` array and GraphQL
query in `taskpane.js`.

## Notes / things to sanity-check before go-live

- I used a generic placeholder icon (blue square, letter K) — swap in real artwork
  in `assets/` if this needs to look polished, and update the manifest GUID
  (`<Id>`) to a fresh one you generate, rather than reusing mine, if this becomes a
  distinct product from other add-ins you might build.
- Batch size (250 IMOs/request) and the `first: 1000` page size are conservative
  defaults — tune in `taskpane.js` if the customer's rate limits allow more, or if
  large batches are timing out.
- No retry/backoff on rate-limit (HTTP 429) responses yet — add if the customer's
  usage volume makes that likely.
