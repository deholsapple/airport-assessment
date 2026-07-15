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

// ---------------------------------------------------------------------
// Task 1 -- build-time field mapping audit.
//
// Every data-mapping loop in this file reads columns off a raw CSV row via
// `r.SOME_COLUMN`. In JS, referencing a field that doesn't exist on `r`
// silently evaluates to `undefined` (then `|| ""` swallows it into an empty
// string) instead of throwing. That is exactly how a previous bug (the
// remarks TAB_NAME filter) hid a 100%-miss field silently. This audit
// statically extracts every `r.FIELD` reference in this file and checks
// that FIELD actually exists in the header of the CSV that loop reads from.
//
// A `// @fields: SOME_FILE.csv` marker comment must sit directly above the
// statement (loop, or filter/map callback) that consumes that table's rows;
// the audit captures that one statement (paren/brace-depth-aware) and
// scans it for r.FIELD references.
// ---------------------------------------------------------------------
function captureStatement(src, start) {
  let i = start;
  while (i < src.length && /\s/.test(src[i])) i++;
  const begin = i;
  let depth = 0;
  let started = false;
  for (; i < src.length; i++) {
    const c = src[i];
    if (c === "(" || c === "{" || c === "[") { depth++; started = true; }
    else if (c === ")" || c === "]") { depth--; }
    else if (c === "}") {
      depth--;
      if (started && depth === 0) { i++; break; }
    } else if (c === ";" && started && depth === 0) { i++; break; }
  }
  return src.slice(begin, i);
}

function auditFieldMappings(sourceText, headersByFile) {
  const markerRe = /\/\/ @fields:\s*([A-Za-z0-9_.]+)\s*\n/g;
  const problems = [];
  const checked = [];
  let m;
  while ((m = markerRe.exec(sourceText))) {
    const file = m[1];
    const header = headersByFile[file];
    if (!header) {
      problems.push(file + ": @fields marker references a CSV that was never loaded");
      continue;
    }
    const stmt = captureStatement(sourceText, markerRe.lastIndex);
    const fieldRe = /\br\.([A-Z][A-Z0-9_]*)\b/g;
    const fields = new Set();
    let fm;
    while ((fm = fieldRe.exec(stmt))) fields.add(fm[1]);
    const missing = [...fields].filter((f) => !header.has(f));
    checked.push({ file, fieldCount: fields.size });
    if (missing.length) {
      problems.push(file + ": referenced column(s) not present in header -> " + missing.join(", "));
    }
  }
  return { problems, checked };
}

// ---------------------------------------------------------------------
// Load every table up front (needed so the audit has all headers before
// any of the mapping code below runs).
// ---------------------------------------------------------------------
const TABLE_FILES = [
  "APT_BASE.csv",
  "APT_RWY.csv",
  "APT_RWY_END.csv",
  "APT_CON.csv",
  "FRQ.csv",
  "AWOS.csv",
  "ATC_BASE.csv",
  "APT_RMK.csv",
  "APT_ATT.csv",
];

const headersByFile = {};
const tables = {};
for (const file of TABLE_FILES) {
  console.log("Loading " + file + " ...");
  const { records, idx } = loadTable(file);
  tables[file] = records;
  headersByFile[file] = new Set(Object.keys(idx));
  console.log("  " + records.length + " rows");
}

const auditResult = auditFieldMappings(fs.readFileSync(__filename, "utf8"), headersByFile);
console.log("\nField mapping audit: checked " + auditResult.checked.length + " mapped section(s) across " +
  new Set(auditResult.checked.map((c) => c.file)).size + " CSV file(s).");
if (auditResult.problems.length) {
  console.error("\nFIELD MAPPING AUDIT FAILED:");
  for (const p of auditResult.problems) console.error("  - " + p);
  console.error("\nRefusing to build with unverified field mappings. Fix the column reference(s) above.");
  process.exit(1);
}
console.log("Field mapping audit: OK -- every r.COLUMN_NAME reference matches its CSV header.\n");

const aptBase = tables["APT_BASE.csv"];
const aptRwy = tables["APT_RWY.csv"];
const aptRwyEnd = tables["APT_RWY_END.csv"];
const aptCon = tables["APT_CON.csv"];
const frq = tables["FRQ.csv"];
const awos = tables["AWOS.csv"];
const atcBase = tables["ATC_BASE.csv"];
const aptRmk = tables["APT_RMK.csv"];
const aptAtt = tables["APT_ATT.csv"];

// ---------------------------------------------------------------------
// Task 2 -- code decode tables.
//
// Sourced from the FAA NASR 28-Day Subscription landing-facility record
// layout (nfdc.faa.gov/webContent/28DaySub/.../Layout_Data/apt_rf.txt),
// which documents the SURFACE_TYPE_CODE / TREATMENT_CODE / RWY_LGT_CODE /
// VGSI_CODE / APCH_LGT_SYSTEM_CODE / FUEL_TYPES domains directly.
//
// Codes not covered below (rare/ambiguous composites) are passed through
// unchanged rather than guessed -- an unrecognized code showing its raw
// form is honest; inventing a label for it is exactly the failure mode
// this rewrite exists to eliminate.
// ---------------------------------------------------------------------
const SURFACE_FULL_OVERRIDES = {
  "OIL&CHIP-T": "Oil & chip, treated",
};
const SURFACE_TOKEN_MAP = {
  ASPH: "Asphalt", CONC: "Concrete", DIRT: "Dirt", TURF: "Turf/grass", GRASS: "Grass",
  GRVL: "Gravel", GRAVEL: "Gravel", SAND: "Sand", SOD: "Sod", SNOW: "Snow", ICE: "Ice",
  WATER: "Water", MATS: "Landing mats", TRTD: "Treated", TREATED: "Treated",
  PSP: "Pierced steel planking", CORAL: "Coral/shell", CALICHE: "Caliche",
  ALUM: "Aluminum", ALUMINUM: "Aluminum", STEEL: "Steel", METAL: "Metal",
  BRICK: "Brick", WOOD: "Wood", DECK: "Deck", "ROOF-TOP": "Rooftop", ROOFTOP: "Rooftop",
};
function decodeSurface(code) {
  if (!code) return "";
  if (SURFACE_FULL_OVERRIDES[code]) return SURFACE_FULL_OVERRIDES[code];
  return code
    .split(/[-/]/)
    .map((t) => SURFACE_TOKEN_MAP[t] || t)
    .join(" / ");
}

const RWY_LIGHT_MAP = {
  HIGH: "High intensity edge lighting (HIRL)",
  MED: "Medium intensity edge lighting (MIRL)",
  LOW: "Low intensity edge lighting (LIRL)",
  NSTD: "Non-standard lighting system",
  PERI: "Perimeter lighting",
  STRB: "Strobe lighting",
  FLD: "Floodlights",
  NONE: "No edge lighting",
};
function decodeRwyLighting(code) {
  if (!code) return "";
  return RWY_LIGHT_MAP[code] || code;
}

const VGSI_MAP = {
  S2L: "2-box SAVASI, left side", S2R: "2-box SAVASI, right side",
  V2L: "2-box VASI, left side", V2R: "2-box VASI, right side",
  V4L: "4-box VASI, left side", V4R: "4-box VASI, right side",
  V6L: "6-box VASI, left side", V6R: "6-box VASI, right side",
  V12: "12-box VASI, both sides", V16: "16-box VASI, both sides",
  P2L: "2-light PAPI, left side", P2R: "2-light PAPI, right side",
  P4L: "4-light PAPI, left side", P4R: "4-light PAPI, right side",
  TRIL: "Tri-color VASI, left side", TRIR: "Tri-color VASI, right side",
  PSIL: "Pulsating/steady-burning VASI, left side", PSIR: "Pulsating/steady-burning VASI, right side",
  PNIL: "System of panels, left side", PNIR: "System of panels, right side",
  NSTD: "Non-standard VASI system", PVT: "Privately owned VGSI",
  VAS: "VASI (unspecified configuration)",
};
function decodeVgsi(code) {
  if (!code) return "";
  return VGSI_MAP[code] || code;
}

const APCH_LGT_MAP = {
  ALSAF: "High-intensity approach lighting, 3,000 ft w/ sequenced flashers (ALSAF)",
  ALSF1: "High-intensity approach lighting w/ sequenced flashers, CAT I (ALSF-1)",
  ALSF2: "High-intensity approach lighting w/ sequenced flashers, CAT II/III (ALSF-2)",
  MALS: "Medium-intensity approach lighting system (MALS)",
  MALSF: "Medium-intensity approach lighting w/ sequenced flashers (MALSF)",
  MALSR: "Medium-intensity approach lighting w/ runway alignment indicator lights (MALSR)",
  SSALS: "Simplified short approach lighting system (SSALS)",
  SSALF: "Simplified short approach lighting w/ sequenced flashers (SSALF)",
  SSALR: "Simplified short approach lighting w/ runway alignment indicator lights (SSALR)",
  ODALS: "Omnidirectional approach lighting system (ODALS)",
  RLLS: "Runway lead-in light system (RLLS)",
  NSTD: "Non-standard approach lighting system",
};
function decodeApchLighting(code) {
  if (!code) return "";
  return APCH_LGT_MAP[code] || code;
}

const FUEL_MAP = {
  "80": "Grade 80 avgas (red)",
  "100": "Grade 100 avgas (green)",
  "100LL": "100LL avgas, low lead (blue)",
  "115": "Grade 115 avgas, military (purple)",
  A: "Jet A",
  "A+": "Jet A w/ icing inhibitor",
  "A++": "Jet A w/ icing inhibitor + conductivity additive",
  "A++10": "Jet A w/ additives + thermal stability enhancer",
  A1: "Jet A-1",
  "A1+": "Jet A-1 w/ icing inhibitor",
  B: "Jet B",
  "B+": "Jet B w/ icing inhibitor",
  J: "Jet fuel (unspecified grade)",
  J4: "JP-4 (military)",
  J5: "JP-5 (military)",
  J8: "JP-8 (military)",
  "J8+10": "JP-8 w/ thermal stability enhancer",
  MOGAS: "Automotive gasoline",
  UL91: "Unleaded 91 octane avgas",
  UL94: "Unleaded 94 octane avgas",
};
function decodeFuelTypes(str) {
  if (!str) return [];
  return str.split(",").map((code) => {
    const c = code.trim();
    return { code: c, label: FUEL_MAP[c] || c };
  });
}

// ---------------------------------------------------------------------
// Airports (APT_BASE.csv)
// ---------------------------------------------------------------------
// Scope: public-use, landing-facility-type "Airport", with an assigned ICAO_ID.
// (This app looks up strictly by 4-letter ICAO identifier, so facilities
// without one can't be reached through it anyway -- this is what keeps the
// embedded dataset a manageable size instead of ~19k nationwide records.)
// @fields: APT_BASE.csv
const included = aptBase.filter(
  (r) => r.SITE_TYPE_CODE === "A" && r.FACILITY_USE_CODE === "PU" && r.ICAO_ID
);
console.log(included.length + " public-use airports with an ICAO_ID (in scope)");

const byArptId = new Map(); // ARPT_ID (FAA LID) -> output record

// @fields: APT_BASE.csv
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
    fuelTypesDesc: decodeFuelTypes(r.FUEL_TYPES),
    lightingSchedule: r.LGT_SKED || "",
    beaconSchedule: r.BCN_LGT_SKED || "",
    towerTypeCode: r.TWR_TYPE_CODE || "",
    customsFlag: r.CUST_FLAG || "N",
    landingFeeFlag: r.LNDG_FEE_FLAG || "",
    far139Type: r.FAR_139_TYPE_CODE || "",
    far139CarrierSvc: r.FAR_139_CARRIER_SER_CODE || "",
    arffCertTypeDate: r.ARFF_CERT_TYPE_DATE || "",
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
    attendance: [],
    weatherPhone: "",
    tower: null,
  };
  byArptId.set(r.ARPT_ID, rec);
}

// ---------------------------------------------------------------------
// Runways (APT_RWY.csv) + runway ends (APT_RWY_END.csv)
// ---------------------------------------------------------------------
console.log("Mapping APT_RWY.csv ...");
const runwayByKey = new Map(); // ARPT_ID|RWY_ID -> runway object

// @fields: APT_RWY.csv
for (const r of aptRwy) {
  const base = byArptId.get(r.ARPT_ID);
  if (!base) continue;
  const rw = {
    id: r.RWY_ID,
    length: r.RWY_LEN ? parseInt(r.RWY_LEN, 10) : null,
    width: r.RWY_WIDTH ? parseInt(r.RWY_WIDTH, 10) : null,
    surface: r.SURFACE_TYPE_CODE || "",
    surfaceDesc: decodeSurface(r.SURFACE_TYPE_CODE),
    condition: r.COND || "", // EXCELLENT/GOOD/FAIR/POOR/FAILED
    treatment: r.TREATMENT_CODE || "",
    pcn: r.PCN || "",
    pavementType: r.PAVEMENT_TYPE_CODE || "",
    lighting: r.RWY_LGT_CODE || "", // HIGH/MED/LOW/FLD/NSTD/PERI/STRB/NONE
    lightingDesc: decodeRwyLighting(r.RWY_LGT_CODE),
    grossWtSW: r.GROSS_WT_SW || "",
    grossWtDW: r.GROSS_WT_DW || "",
    grossWtDTW: r.GROSS_WT_DTW || "",
    grossWtDDTW: r.GROSS_WT_DDTW || "",
    ends: [],
  };
  base.runways.push(rw);
  runwayByKey.set(r.ARPT_ID + "|" + r.RWY_ID, rw);
}

console.log("Mapping APT_RWY_END.csv ...");
// @fields: APT_RWY_END.csv
for (const r of aptRwyEnd) {
  const rw = runwayByKey.get(r.ARPT_ID + "|" + r.RWY_ID);
  if (!rw) continue;
  const displacedThrLen = r.DISPLACED_THR_LEN ? parseInt(r.DISPLACED_THR_LEN, 10) : null;
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
    vgsiDesc: decodeVgsi(r.VGSI_CODE),
    apchLgtSystem: r.APCH_LGT_SYSTEM_CODE || "",
    apchLgtSystemDesc: decodeApchLighting(r.APCH_LGT_SYSTEM_CODE),
    reilFlag: r.RWY_END_LGTS_FLAG || "",
    centerlineLgtFlag: r.CNTRLN_LGTS_AVBL_FLAG || "",
    tdzLgtFlag: r.TDZ_LGT_AVBL_FLAG || "",
    displacedThrLen: r.DISPLACED_THR_LEN || "",
    hasDisplacedThr: !!(displacedThrLen && displacedThrLen > 0),
    tch: r.THR_CROSSING_HGT || "",
    glidePathAngle: r.VISUAL_GLIDE_PATH_ANGLE || "",
    elev: r.RWY_END_ELEV || "",
    // Declared distances (FAR 1.2 / TERPS terms). These are per-END and
    // directional -- a displaced threshold or an obstacle can make the
    // declared distance shorter than the runway's physical length, and
    // that shortfall only shows up on the affected end/direction. Using
    // physical length in place of a missing declared distance would
    // over-credit the runway, which is the wrong direction to be wrong
    // in. Blank here means NASR has not published one for this end --
    // that must surface as NOT AVAILABLE, never silently fall back to
    // physical length.
    tora: r.TKOF_RUN_AVBL || "", // Takeoff Run Available
    toda: r.TKOF_DIST_AVBL || "", // Takeoff Distance Available
    asda: r.ACLT_STOP_DIST_AVBL || "", // Accelerate-Stop Distance Available
    lda: r.LNDG_DIST_AVBL || "", // Landing Distance Available
  });
}

// ---------------------------------------------------------------------
// Contacts (APT_CON.csv)
// ---------------------------------------------------------------------
console.log("Mapping APT_CON.csv ...");
// @fields: APT_CON.csv
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
//
// BUG FIX: this used to hard-cap at the first 20 rows per airport. FRQ.csv
// lists an airport's TRACON-served approach/departure procedure frequencies
// (one row per named STAR/DP) before its own tower/ground/CTAF/ATIS rows,
// so a busy Class-B-served field could burn through the cap on procedure
// entries alone and silently lose its tower and CTAF frequency. There is no
// cap now -- every serviced-facility row is kept. They're sorted so the
// frequencies a pilot actually wants first (tower, CTAF, weather) sort
// ahead of the bulk STAR/DP list, which is cosmetic only, not a filter.
// ---------------------------------------------------------------------
console.log("Mapping FRQ.csv ...");
const FREQ_TYPE_PRIORITY = [
  "ATCT", "NON-ATCT", "ASOS_AWOS", "FSS", "NAVAID", "RCO", "RCO1", "RCAG", "TRACON", "ARTCC", "CERAP",
];
// @fields: FRQ.csv
for (const r of frq) {
  const base = byArptId.get(r.SERVICED_FACILITY);
  if (!base) continue;
  base.frequencies.push({
    facilityType: r.FACILITY_TYPE || "",
    call: r.TOWER_OR_COMM_CALL || r.PRIMARY_APPROACH_RADIO_CALL || "",
    freq: r.FREQ || "",
    use: r.FREQ_USE || "",
    remark: (r.REMARK || "").slice(0, 300),
  });
}
for (const base of byArptId.values()) {
  if (base.frequencies.length < 2) continue;
  base.frequencies.sort((a, b) => {
    const pa = FREQ_TYPE_PRIORITY.indexOf(a.facilityType);
    const pb = FREQ_TYPE_PRIORITY.indexOf(b.facilityType);
    return (pa === -1 ? 999 : pa) - (pb === -1 ? 999 : pb);
  });
}

// ---------------------------------------------------------------------
// AWOS/ASOS weather station phone (AWOS.csv)
// ---------------------------------------------------------------------
console.log("Mapping AWOS.csv ...");
// AWOS.csv doesn't carry ARPT_ID directly; join via ASOS_AWOS_ID which
// matches the FAA LID (ARPT_ID) in the common case of an on-airport station.
// @fields: AWOS.csv
for (const r of awos) {
  const base = byArptId.get(r.ASOS_AWOS_ID);
  if (!base) continue;
  if (!base.weatherPhone && r.PHONE_NO) base.weatherPhone = r.PHONE_NO;
}

// ---------------------------------------------------------------------
// Tower info (ATC_BASE.csv) -- only present for towered airports
//
// BUG FIX: ATC_BASE.csv also carries NON-ATCT rows (and ARTCC/TRACON/CERAP
// area-facility rows) -- it is not "one row per tower". The old code
// attached a tower object to every row it could join, which reported 1,707
// uncontrolled airports as towered. Only FACILITY_TYPE values that start
// with "ATCT" (ATCT, ATCT-A/C, ATCT-RAPCON, ATCT-RATCF, ATCT-TRACON) denote
// an actual staffed tower at that airport.
// ---------------------------------------------------------------------
console.log("Mapping ATC_BASE.csv ...");
// @fields: ATC_BASE.csv
for (const r of atcBase) {
  if (!r.FACILITY_TYPE || r.FACILITY_TYPE.indexOf("ATCT") !== 0) continue;
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
//
// BUG FIX: TAB_NAME in this CSV holds NASR's table names (AIRPORT, RUNWAY,
// RUNWAY_END, RUNWAY_END_OBSTN, AIRPORT_ATTEND_SCHED, ...), not the CSV
// filenames those tables ship as (APT_BASE.csv etc). The old filter set
// compared against filename-shaped strings that never occur in this
// column, so it matched zero rows and every airport lost every remark.
// ---------------------------------------------------------------------
console.log("Mapping APT_RMK.csv ...");
const RELEVANT_TABS = new Set(["AIRPORT", "RUNWAY", "RUNWAY_END"]);
// @fields: APT_RMK.csv
for (const r of aptRmk) {
  if (!RELEVANT_TABS.has(r.TAB_NAME)) continue;
  const base = byArptId.get(r.ARPT_ID);
  if (!base) continue;
  if (base.remarks.length >= 25) continue; // sane cap
  const tag = r.REF_COL_NAME ? r.REF_COL_NAME + ": " : "";
  base.remarks.push((tag + r.REMARK).slice(0, 400));
}

// ---------------------------------------------------------------------
// Attendance / operating hours (APT_ATT.csv)
//
// Previously not loaded at all -- the only hint of "when is this airport
// staffed" available downstream was LGT_SKED (the *lighting* schedule),
// which answers a different question (when do the lights turn on, often
// dusk-dawn regardless of staffing) and must not be read as attendance.
// APT_ATT carries the actual attendance schedule as one row per
// month/day-group segment (e.g. ALL/MON-FRI/0800-1800 + ALL/SAT/0900-1700);
// an airport with no rows here simply has no attendance data published.
// ---------------------------------------------------------------------
console.log("Mapping APT_ATT.csv ...");
// @fields: APT_ATT.csv
for (const r of aptAtt) {
  const base = byArptId.get(r.ARPT_ID);
  if (!base) continue;
  base.attendance.push({
    seq: r.SKED_SEQ_NO || "",
    month: r.MONTH || "",
    days: r.DAY || "",
    hours: r.HOUR || "",
  });
}

function titleCase(s) {
  if (!s) return "";
  return s.replace(/\w\S*/g, (t) => t.charAt(0) + t.slice(1).toLowerCase());
}

// ---------------------------------------------------------------------
// Emit
// ---------------------------------------------------------------------
const out = {};
for (const [, rec] of byArptId.entries()) {
  out[rec.icao] = rec;
}

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
