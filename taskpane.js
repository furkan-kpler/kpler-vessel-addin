/* global Office, Excel */

const STORAGE_KEYS = {
  endpoint: "kplerVesselAddin.endpoint",
  apiKey: "kplerVesselAddin.apiKey",
  authScheme: "kplerVesselAddin.authScheme",
};

const BATCH_SIZE = 250; // IMOs per GraphQL request (well under the 1000 "first" cap)

let els = {};

Office.onReady(() => {
  els = {
    endpoint: document.getElementById("endpoint-input"),
    apiKey: document.getElementById("apikey-input"),
    authScheme: document.getElementById("auth-scheme-select"),
    saveSettingsBtn: document.getElementById("save-settings-btn"),
    settingsDetails: document.getElementById("settings-details"),
    imoInput: document.getElementById("imo-input"),
    sheetName: document.getElementById("sheet-name-input"),
    fetchBtn: document.getElementById("fetch-btn"),
    statusText: document.getElementById("status-text"),
    progressBar: document.getElementById("progress-bar"),
    logList: document.getElementById("log-list"),
  };

  loadSettings();
  els.saveSettingsBtn.addEventListener("click", saveSettings);
  els.fetchBtn.addEventListener("click", onFetchClick);

  // Open the settings panel automatically if nothing is configured yet
  if (!els.endpoint.value) {
    els.settingsDetails.setAttribute("open", "open");
  }
});

function loadSettings() {
  els.endpoint.value = localStorage.getItem(STORAGE_KEYS.endpoint) || "";
  els.apiKey.value = localStorage.getItem(STORAGE_KEYS.apiKey) || "";
  els.authScheme.value = localStorage.getItem(STORAGE_KEYS.authScheme) || "Basic";
}

function saveSettings() {
  localStorage.setItem(STORAGE_KEYS.endpoint, els.endpoint.value.trim());
  localStorage.setItem(STORAGE_KEYS.apiKey, els.apiKey.value.trim());
  localStorage.setItem(STORAGE_KEYS.authScheme, els.authScheme.value);
  setStatus("Settings saved in this browser.", "success");
  els.settingsDetails.removeAttribute("open");
}

function setStatus(msg, kind) {
  els.statusText.textContent = msg;
  els.statusText.className = kind || "";
}

function log(msg) {
  const li = document.createElement("li");
  li.textContent = msg;
  els.logList.prepend(li);
}

function buildAuthHeaders() {
  const key = els.apiKey.value.trim();
  const scheme = els.authScheme.value;
  const headers = { "Content-Type": "application/json" };
  if (!key) return headers;
  if (scheme === "Basic") {
    // Pass the token through exactly as given - whether the customer pasted just the
    // token or the whole "Basic <token>" string, either way works, no re-encoding.
    const alreadyPrefixed = /^basic\s+/i.test(key);
    headers["Authorization"] = alreadyPrefixed ? key : `Basic ${key}`;
  } else if (scheme === "Bearer") headers["Authorization"] = `Bearer ${key}`;
  else if (scheme === "Raw") headers["Authorization"] = key;
  else if (scheme === "ApiKey") headers["x-api-key"] = key;
  return headers;
}

function parseImos(raw) {
  return raw
    .split(/[\s,;]+/)
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
}

function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

const GRAPHQL_QUERY = `
query VesselLookup($imos: [String!]!, $after: String) {
  vessels(first: 1000, after: $after, where: { filters: [{ field: "identifier.imo", op: IN, values: $imos }] }) {
    nodes {
      identifier { imo mmsi callSign eni shipId }
      particulars {
        general { name commercialFleet generalVesselType detailedVesselType serviceStatus flag portOfRegistry keelLaidDate }
        hull { yearOfBuild yardNumber hullMaterial hullType decks }
        dimension { lengthOverall lengthBetweenPerpendiculars breadthExtreme breadthMoulded draught depth freeboard }
        tonnage { grossTonnage deadweightTonnage netTonnage loadedDisplacementTonnage lightDisplacementTonnage }
        capacity { liquidCapacity gasCapacity baleCapacity grainCapacity teuCapacity ceuCapacity passengerCapacity ballastCapacity }
        engine { enginePower engineUnits engineCylinderUnits engineBore engineStroke engineRpm engineType speedService propeller }
        fuel { mainEngineFuelType fuelCapacity }
        aisTransceiver { lengthFore lengthAft widthLeft widthRight aisTransceiverClass }
      }
      ownership {
        ultimateBeneficialOwner { entityType name countryOfResidence startDate confidenceLevel shareholderPercentage }
        beneficialOwner { current { name country address website email startDate endDate } }
        registeredOwner { current { name country address website email startDate endDate } }
      }
      commercial {
        disponentOwner { current { name country startDate endDate } }
        commercialOperator { current { name country startDate endDate } }
      }
      management {
        technicalManager { current { name country startDate endDate } }
        ismManager { current { name country startDate endDate } }
      }
      associatedCompanies {
        shipBuilder { current { name country } }
        engineBuilder { current { name country } }
        classificationSociety { current { name country } }
        piClub { current { name country } }
      }
      historical { names { name endDate } }
    }
    pageInfo { hasNextPage endCursor }
  }
}`;

async function fetchBatch(imos, after) {
  const res = await fetch(els.endpoint.value.trim(), {
    method: "POST",
    headers: buildAuthHeaders(),
    body: JSON.stringify({ query: GRAPHQL_QUERY, variables: { imos, after: after || null } }),
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`HTTP ${res.status} ${res.statusText}${text ? " - " + text.slice(0, 300) : ""}`);
  }

  const json = await res.json();
  if (json.errors && json.errors.length) {
    throw new Error(json.errors.map((e) => e.message).join("; "));
  }
  return json.data.vessels;
}

async function fetchAllVessels(imoList, onProgress) {
  const batches = chunk(imoList, BATCH_SIZE);
  const allNodes = [];

  for (let i = 0; i < batches.length; i++) {
    let after = null;
    let more = true;
    while (more) {
      const page = await fetchBatch(batches[i], after);
      allNodes.push(...page.nodes);
      more = page.pageInfo.hasNextPage;
      after = page.pageInfo.endCursor;
    }
    onProgress((i + 1) / batches.length);
  }
  return allNodes;
}

// ---- Flattening: one row per vessel, one column per field ----

function fmtCompany(c) {
  if (!c) return "";
  const parts = [c.name, c.country].filter(Boolean);
  return parts.join(" / ");
}

function fmtUbo(list) {
  if (!list || !list.length) return "";
  return list
    .map((u) => {
      const pct = u.shareholderPercentage != null ? `${u.shareholderPercentage}%` : "";
      return [u.name, u.countryOfResidence, pct, u.confidenceLevel].filter(Boolean).join(" / ");
    })
    .join("; ");
}

function fmtHistoricalNames(h) {
  if (!h || !h.names || !h.names.length) return "";
  return h.names.map((n) => `${n.name} (until ${n.endDate})`).join("; ");
}

const COLUMNS = [
  ["IMO", (n) => n.identifier.imo],
  ["MMSI", (n) => n.identifier.mmsi],
  ["Call Sign", (n) => n.identifier.callSign],
  ["ENI", (n) => n.identifier.eni],
  ["Ship ID", (n) => n.identifier.shipId],
  ["Name", (n) => n.particulars.general.name],
  ["Commercial Fleet", (n) => n.particulars.general.commercialFleet],
  ["General Vessel Type", (n) => n.particulars.general.generalVesselType],
  ["Detailed Vessel Type", (n) => n.particulars.general.detailedVesselType],
  ["Service Status", (n) => n.particulars.general.serviceStatus],
  ["Flag", (n) => n.particulars.general.flag],
  ["Port of Registry", (n) => n.particulars.general.portOfRegistry],
  ["Keel Laid Date", (n) => n.particulars.general.keelLaidDate],
  ["Year of Build", (n) => n.particulars.hull.yearOfBuild],
  ["Yard Number", (n) => n.particulars.hull.yardNumber],
  ["Hull Material", (n) => n.particulars.hull.hullMaterial],
  ["Hull Type", (n) => n.particulars.hull.hullType],
  ["Decks", (n) => n.particulars.hull.decks],
  ["Length Overall (m)", (n) => n.particulars.dimension.lengthOverall],
  ["Length Between Perpendiculars (m)", (n) => n.particulars.dimension.lengthBetweenPerpendiculars],
  ["Breadth Extreme (m)", (n) => n.particulars.dimension.breadthExtreme],
  ["Breadth Moulded (m)", (n) => n.particulars.dimension.breadthMoulded],
  ["Draught (m)", (n) => n.particulars.dimension.draught],
  ["Depth (m)", (n) => n.particulars.dimension.depth],
  ["Freeboard (mm)", (n) => n.particulars.dimension.freeboard],
  ["Gross Tonnage", (n) => n.particulars.tonnage.grossTonnage],
  ["Deadweight Tonnage", (n) => n.particulars.tonnage.deadweightTonnage],
  ["Net Tonnage", (n) => n.particulars.tonnage.netTonnage],
  ["Loaded Displacement Tonnage", (n) => n.particulars.tonnage.loadedDisplacementTonnage],
  ["Light Displacement Tonnage", (n) => n.particulars.tonnage.lightDisplacementTonnage],
  ["Liquid Capacity (m3)", (n) => n.particulars.capacity.liquidCapacity],
  ["Gas Capacity (m3)", (n) => n.particulars.capacity.gasCapacity],
  ["Bale Capacity (m3)", (n) => n.particulars.capacity.baleCapacity],
  ["Grain Capacity (m3)", (n) => n.particulars.capacity.grainCapacity],
  ["TEU Capacity", (n) => n.particulars.capacity.teuCapacity],
  ["CEU Capacity", (n) => n.particulars.capacity.ceuCapacity],
  ["Passenger Capacity", (n) => n.particulars.capacity.passengerCapacity],
  ["Ballast Capacity (m3)", (n) => n.particulars.capacity.ballastCapacity],
  ["Engine Power (kW)", (n) => n.particulars.engine.enginePower],
  ["Engine Units", (n) => n.particulars.engine.engineUnits],
  ["Engine Cylinder Units", (n) => n.particulars.engine.engineCylinderUnits],
  ["Engine Bore (mm)", (n) => n.particulars.engine.engineBore],
  ["Engine Stroke (mm)", (n) => n.particulars.engine.engineStroke],
  ["Engine RPM", (n) => n.particulars.engine.engineRpm],
  ["Engine Type", (n) => n.particulars.engine.engineType],
  ["Service Speed (kn)", (n) => n.particulars.engine.speedService],
  ["Propeller", (n) => n.particulars.engine.propeller],
  ["Main Engine Fuel Type", (n) => n.particulars.fuel.mainEngineFuelType],
  ["Fuel Capacity (m3)", (n) => n.particulars.fuel.fuelCapacity],
  ["AIS Transceiver Class", (n) => n.particulars.aisTransceiver.aisTransceiverClass],
  ["Ultimate Beneficial Owner(s)", (n) => fmtUbo(n.ownership.ultimateBeneficialOwner)],
  ["Beneficial Owner (current)", (n) => fmtCompany(n.ownership.beneficialOwner.current)],
  ["Registered Owner (current)", (n) => fmtCompany(n.ownership.registeredOwner.current)],
  ["Disponent Owner (current)", (n) => fmtCompany(n.commercial.disponentOwner.current)],
  ["Commercial Operator (current)", (n) => fmtCompany(n.commercial.commercialOperator.current)],
  ["Technical Manager (current)", (n) => fmtCompany(n.management.technicalManager.current)],
  ["ISM Manager (current)", (n) => fmtCompany(n.management.ismManager.current)],
  ["Ship Builder", (n) => fmtCompany(n.associatedCompanies.shipBuilder.current)],
  ["Engine Builder", (n) => fmtCompany(n.associatedCompanies.engineBuilder.current)],
  ["Classification Society", (n) => fmtCompany(n.associatedCompanies.classificationSociety.current)],
  ["P&I Club", (n) => fmtCompany(n.associatedCompanies.piClub.current)],
  ["Historical Names", (n) => fmtHistoricalNames(n.historical)],
];

function toRow(node) {
  return COLUMNS.map(([, get]) => {
    try {
      const v = get(node);
      return v === null || v === undefined ? "" : v;
    } catch (e) {
      return "";
    }
  });
}

async function writeToSheet(nodes) {
  const sheetName = (els.sheetName.value || "Vessel Data").trim() || "Vessel Data";
  const headers = COLUMNS.map(([h]) => h);
  const rows = nodes.map(toRow);

  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    let sheet = sheets.items.find((s) => s.name === sheetName);
    if (sheet) {
      sheet.getRange().clear();
    } else {
      sheet = sheets.add(sheetName);
    }

    const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
    headerRange.values = [headers];

    if (rows.length > 0) {
      const dataRange = sheet.getRangeByIndexes(1, 0, rows.length, headers.length);
      dataRange.values = rows;
    }

    const usedRange = sheet.getRangeByIndexes(0, 0, rows.length + 1, headers.length);
    let table;
    try {
      table = sheet.tables.add(usedRange, true);
      table.name = sheetName.replace(/[^A-Za-z0-9_]/g, "_") + "Tbl";
    } catch (e) {
      // A table may already exist over this range from a previous run - ignore.
    }

    usedRange.format.autofitColumns();
    sheet.activate();
    await context.sync();
  });
}

async function onFetchClick() {
  const endpoint = els.endpoint.value.trim();
  if (!endpoint) {
    setStatus("Set the GraphQL endpoint URL in Connection settings first.", "error");
    els.settingsDetails.setAttribute("open", "open");
    return;
  }

  const imoList = parseImos(els.imoInput.value);
  if (imoList.length === 0) {
    setStatus("Paste at least one IMO number.", "error");
    return;
  }

  els.fetchBtn.disabled = true;
  els.progressBar.style.display = "block";
  els.progressBar.value = 0;
  setStatus(`Looking up ${imoList.length} vessel(s)...`);
  log(`Requested ${imoList.length} IMO(s).`);

  try {
    const nodes = await fetchAllVessels(imoList, (frac) => {
      els.progressBar.value = Math.round(frac * 70);
    });
    log(`Received ${nodes.length} vessel record(s) from Kpler.`);

    if (nodes.length < imoList.length) {
      const foundImos = new Set(nodes.map((n) => String(n.identifier.imo)));
      const missing = imoList.filter((imo) => !foundImos.has(String(imo)));
      if (missing.length) log(`No match for: ${missing.join(", ")}`);
    }

    setStatus("Writing results to the worksheet...");
    els.progressBar.value = 85;
    await writeToSheet(nodes);
    els.progressBar.value = 100;

    setStatus(`Done - wrote ${nodes.length} vessel record(s).`, "success");
  } catch (err) {
    console.error(err);
    setStatus(`Error: ${err.message}`, "error");
    log(`Error: ${err.message}`);
  } finally {
    els.fetchBtn.disabled = false;
    setTimeout(() => {
      els.progressBar.style.display = "none";
    }, 800);
  }
}
