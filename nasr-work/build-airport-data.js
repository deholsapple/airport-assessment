// One-time / periodic preprocessing script.
// Downloads are NOT done here -- run after extracting the FAA 28-Day NASR CSV
// subscription (CSV_Data/<cycle>_CSV.zip) into ./extracted/CSV_Data/csv/
//
// Usage: node build-airport-data.js
// Output: ../airport-data.js  (window.NASR_AIRPORT_DATA = {...}, keyed by ICAO id)

const fs = require("fs");
const path = require("path");

const CSV_DIR = path.join(__dirname, "extracted", "CSV_Data", "csv");
const OUT_FILE = path.join(__dirname, "..", "airport-data.js");

// ---------------------------------------------------------------------
// Minimal RFC4180-ish CSV parser (handles quoted fields, "" escaping,
// embedded commas and embedded newlines inside quoted fields).
// ---------------------------------------------------------------------
function parseCSV(text) {
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;
  let i = 0;
  const len = text.length;
  while (i < len) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
        inQuotes = false; i++; continue;
      }
      field += c; i++; continue;
    } else {
      if (c === '"') { inQuotes = true; i++; continue; }
      if (c === ',') { row.push(field); field = ""; i++; continue; }
      if (c === '\r') { i++; continue; }
      if (c === '\n') { row.push(field); rows.push(row); row = []; field = ""; i++; continue; }
      field += c; i++; continue;
    }
  }
  if (field.length || row.length) { row.push(field); rows.push(row); }
  return rows;
}

function loadTable(filename) {
  const text = fs.readFileSync(path.join(CSV_DIR, filename), "utf8");
  const rows = parseCSV(text).filter((r) => r.length > 1);
  const header = rows[0];
  const idx = {};
  header.forEach((h, i) => { idx[h] = i; });
  const records = rows.slice(1).map((r) => {
    const o = {};
    header.forEach((h, i) => { o[h] = (r[i] || "").trim(); });
    return o;
  });
  return { records, idx };
}

console.log("Loading APT_BASE.csv ...");
const aptBase = loadTable("APT_BASE.csv").records;
console.log("  " + aptBase.length + " rows");

// Scope: public-use, landing-facility-type "Airport", with an assigned ICAO_ID.
// (This app looks up strictly by 4-letter ICAO identifier, so facilities
// without one can't be reached through it anyway -- this is what keeps the
// embedded dataset a manageable size instead of ~19k nationwide records.)
const included = aptBase.filter(
  (r) => r.SITE_TYPE_CODE === "A" && r.FACILITY_USE_CODE === "PU" && r.ICAO_ID
);
console.log("  " + included.length + " public-use airports with an ICAO_ID (in scope)");

const byArptId = new Map(); // ARPT_ID (FAA LID) -> output record
const arptIdToIcao = new Map();

for (const r of included) {
  const rec = {
    icao: r.ICAO_ID,
    faaId: r.ARPT_ID,
    name: titleCase(r.ARPT_NAME),
    city: titleCase(r.CITY),
    state: r.STATE_CODE,
    county: titleCase(r.COUNTY_NAME),
    lat: parseFloat(r.LAT_DECIMAL) || null,
    lon: parseFloat(r.LONG_DECIMAL) || null,
    elev: r.ELEV ? parseFloat(r.ELEV) : null,
    magVarn: r.MAG_VARN ? r.MAG_VARN + (r.MAG_HEMIS || "") : "",
    tpa: r.TPA || "",
    ownership: r.OWNERSHIP_TYPE_CODE, // PU/PR/MA/MN/MR/CG
    facilityUse: r.FACILITY_USE_CODE, // PU/PR
    status: r.ARPT_STATUS, // O / CI / CP
    fuelTypes: r.FUEL_TYPES || "",
    lightingSchedule: r.LGT_SKED || "",
    beaconSchedule: r.BCN_LGT_SKED || "",
    towerTypeCode: r.TWR_TYPE_CODE || "",
    customsFlag: r.CUST_FLAG || "N",
    landingFeeFlag: r.LNDG_FEE_FLAG || "",
    far139Type: r.FAR_139_TYPE_CODE || "",
    far139CarrierSvc: r.FAR_139_CARRIER_SER_CODE || "",
    notamId: r.NOTAM_ID || "",
    notamFlag: r.NOTAM_FLAG || "",
    fssId: r.FSS_ID || "",
    fssName: titleCase(r.FSS_NAME || ""),
    fssPhone: r.PHONE_NO || "",
    fssTollFree: r.TOLL_FREE_NO || "",
    distCityToAirport: r.DIST_CITY_TO_AIRPORT || "",
    directionCode: r.DIRECTION_CODE || "",
    effDate: r.EFF_DATE || "",
    runways: [],
    frequencies: [],
    contacts: [],
    remarks: [],
    weatherPhone: "",
    tower: null,
  };
  byArptId.set(r.ARPT_ID, rec);
  arptIdToIcao.set(r.ARPT_ID, r.ICAO_ID);
}

// ---------------------------------------------------------------------
// Runways (APT_RWY.csv) + runway ends (APT_RWY_END.csv)
// ---------------------------------------------------------------------
console.log("Loading APT_RWY.csv ...");
const aptRwy = loadTable("APT_RWY.csv").records;
const runwayByKey = new Map(); // ARPT_ID|RWY_ID -> runway object

for (const r of aptRwy) {
  const base = byArptId.get(r.ARPT_ID);
  if (!base) continue;
  const rw = {
    id: r.RWY_ID,
    length: r.RWY_LEN ? parseInt(r.RWY_LEN, 10) : null,
    width: r.RWY_WIDTH ? parseInt(r.RWY_WIDTH, 10) : null,
    surface: r.SURFACE_TYPE_CODE || "",
    condition: r.COND || "", // EXCELLENT/GOOD/FAIR/POOR/FAILED
    treatment: r.TREATMENT_CODE || "",
    pcn: r.PCN || "",
    pavementType: r.PAVEMENT_TYPE_CODE || "",
    lighting: r.RWY_LGT_CODE || "", // HIGH/MED/LOW/FLD/NSTD/PERI/STRB/NONE
    grossWtSW: r.GROSS_WT_SW || "",
    grossWtDW: r.GROSS_WT_DW || "",
    grossWtDTW: r.GROSS_WT_DTW || "",
    grossWtDDTW: r.GROSS_WT_DDTW || "",
    ends: [],
  };
  base.runways.push(rw);
  runwayByKey.set(r.ARPT_ID + "|" + r.RWY_ID, rw);
}

console.log("Loading APT_RWY_END.csv ...");
const aptRwyEnd = loadTable("APT_RWY_END.csv").records;
for (const r of aptRwyEnd) {
  const rw = runwayByKey.get(r.ARPT_ID + "|" + r.RWY_ID);
  if (!rw) continue;
  rw.ends.push({
    id: r.RWY_END_ID,
    trueAlignment: r.TRUE_ALIGNMENT || "",
    // Ground-based nav aid at this end, if any. NOTE: this does NOT cover
    // RNAV(GPS)-only approaches -- those live in separate FAA TPP chart
    // data, not this dataset. Absence here must never be read as "no
    // instrument approach" / "VFR only".
    ilsType: r.ILS_TYPE || "",
    markingType: r.RWY_MARKING_TYPE_CODE || "",
    markingCond: r.RWY_MARKING_COND || "", // G/F/P
    vgsi: r.VGSI_CODE || "",
    apchLgtSystem: r.APCH_LGT_SYSTEM_CODE || "",
    reilFlag: r.RWY_END_LGTS_FLAG || "",
    centerlineLgtFlag: r.CNTRLN_LGTS_AVBL_FLAG || "",
    tdzLgtFlag: r.TDZ_LGT_AVBL_FLAG || "",
    displacedThrLen: r.DISPLACED_THR_LEN || "",
    tch: r.THR_CROSSING_HGT || "",
    glidePathAngle: r.VISUAL_GLIDE_PATH_ANGLE || "",
    elev: r.RWY_END_ELEV || "",
  });
}

// ---------------------------------------------------------------------
// Contacts (APT_CON.csv)
// ---------------------------------------------------------------------
console.log("Loading APT_CON.csv ...");
const aptCon = loadTable("APT_CON.csv").records;
for (const r of aptCon) {
  const base = byArptId.get(r.ARPT_ID);
  if (!base) continue;
  base.contacts.push({
    title: titleCase(r.TITLE || ""),
    name: titleCase(r.NAME || ""),
    phone: r.PHONE_NO || "",
    address: [r.ADDRESS1, r.ADDRESS2].filter(Boolean).join(", "),
  });
}

// ---------------------------------------------------------------------
// Frequencies (FRQ.csv) -- CTAF/UNICOM/ASOS-AWOS/RCO/RCAG/tower etc.
// ---------------------------------------------------------------------
console.log("Loading FRQ.csv ...");
const frq = loadTable("FRQ.csv").records;
for (const r of frq) {
  const base = byArptId.get(r.SERVICED_FACILITY);
  if (!base) continue;
  if (base.frequencies.length >= 20) continue; // sane cap
  base.frequencies.push({
    facilityType: r.FACILITY_TYPE || "",
    call: r.TOWER_OR_COMM_CALL || r.PRIMARY_APPROACH_RADIO_CALL || "",
    freq: r.FREQ || "",
    use: r.FREQ_USE || "",
    remark: (r.REMARK || "").slice(0, 300),
  });
}

// ---------------------------------------------------------------------
// AWOS/ASOS weather station phone (AWOS.csv)
// ---------------------------------------------------------------------
console.log("Loading AWOS.csv ...");
const awos = loadTable("AWOS.csv").records;
// AWOS.csv doesn't carry ARPT_ID directly; join via ASOS_AWOS_ID which
// matches the FAA LID (ARPT_ID) in the common case of an on-airport station.
for (const r of awos) {
  const base = byArptId.get(r.ASOS_AWOS_ID);
  if (!base) continue;
  if (!base.weatherPhone && r.PHONE_NO) base.weatherPhone = r.PHONE_NO;
}

// ---------------------------------------------------------------------
// Tower info (ATC_BASE.csv) -- only present for towered airports
// ---------------------------------------------------------------------
console.log("Loading ATC_BASE.csv ...");
const atcBase = loadTable("ATC_BASE.csv").records;
for (const r of atcBase) {
  const base = byArptId.get(r.FACILITY_ID);
  if (!base) continue;
  base.tower = {
    call: r.TWR_CALL || "",
    hours: r.TWR_HRS || "",
    apchProvider: r.APCH_P_PROVIDER || "",
    depProvider: r.DEP_P_PROVIDER || "",
  };
}

// ---------------------------------------------------------------------
// Remarks (APT_RMK.csv) -- scoped to the tables we actually surface, capped.
// ---------------------------------------------------------------------
console.log("Loading APT_RMK.csv ...");
const aptRmk = loadTable("APT_RMK.csv").records;
const RELEVANT_TABS = new Set(["APT_BASE", "APT_RWY", "APT_RWY_END"]);
for (const r of aptRmk) {
  if (!RELEVANT_TABS.has(r.TAB_NAME)) continue;
  const base = byArptId.get(r.ARPT_ID);
  if (!base) continue;
  if (base.remarks.length >= 25) continue; // sane cap
  const tag = r.REF_COL_NAME ? r.REF_COL_NAME + ": " : "";
  base.remarks.push((tag + r.REMARK).slice(0, 400));
}

function titleCase(s) {
  if (!s) return "";
  return s.replace(/\w\S*/g, (t) => t.charAt(0) + t.slice(1).toLowerCase());
}

// ---------------------------------------------------------------------
// Emit
// ---------------------------------------------------------------------
const out = {};
for (const [arptId, rec] of byArptId.entries()) {
  out[rec.icao] = rec;
}

const cycleMatch = fs.readFileSync(path.join(CSV_DIR, "APT_BASE.csv"), "utf8").slice(0, 40);
const effDateGuess = (included[0] && included[0].EFF_DATE) || "unknown";

const header =
  "// Auto-generated from the FAA 28-Day NASR CSV Subscription.\n" +
  "// Effective date: " + effDateGuess + "\n" +
  "// Generated: " + new Date().toISOString() + "\n" +
  "// Source: https://nfdc.faa.gov/webContent/28DaySub/ (public domain, U.S. Government work)\n" +
  "// Scope: public-use airports (FACILITY_USE_CODE=PU, SITE_TYPE_CODE=A) with an assigned ICAO_ID.\n" +
  "// Regenerate periodically with build-airport-data.js as the FAA cycle updates.\n" +
  "// IMPORTANT: still verify safety-critical fields against the current Chart Supplement/NOTAMs.\n";

const json = JSON.stringify(out);
fs.writeFileSync(OUT_FILE, header + "window.NASR_AIRPORT_DATA = " + json + ";\n");

console.log("Wrote " + Object.keys(out).length + " airports to " + OUT_FILE);
console.log("Output size: " + (fs.statSync(OUT_FILE).size / 1024 / 1024).toFixed(2) + " MB");
