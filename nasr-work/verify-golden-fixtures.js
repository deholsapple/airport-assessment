// Task 3 -- golden fixtures.
//
// Five real airports, hand-verified against the source FAA NASR CSVs
// (2026/06/11 cycle) field-by-field. These are the values the build MUST
// reproduce on every run; a diff here means a regression, not a data
// update, since the underlying NASR facts for these specific fields are
// stable across cycles for these airports.
//
// Usage: node build-airport-data.js && node verify-golden-fixtures.js

const fs = require("fs");
const path = require("path");
const vm = require("vm");

const OUT_FILE = path.join(__dirname, "..", "airport-data.js");
const src = fs.readFileSync(OUT_FILE, "utf8");
const sandbox = { window: {} };
vm.createContext(sandbox);
vm.runInContext(src, sandbox, { filename: OUT_FILE });
const DATA = sandbox.window.NASR_AIRPORT_DATA;

let failures = 0;
let checks = 0;

function eq(desc, actual, expected) {
  checks++;
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a !== e) {
    failures++;
    console.error("FAIL: " + desc);
    console.error("  expected: " + e);
    console.error("  actual:   " + a);
  } else {
    console.log("ok - " + desc);
  }
}

function ok(desc, cond) {
  checks++;
  if (!cond) {
    failures++;
    console.error("FAIL: " + desc);
  } else {
    console.log("ok - " + desc);
  }
}

function findRunway(rec, id) {
  return rec.runways.find((r) => r.id === id);
}
function findEnd(rw, id) {
  return rw.ends.find((e) => e.id === id);
}

// -----------------------------------------------------------------------
// KDTO -- Denton Enterprise, TX. Towered. Confirms: bug #1 (frequency cap
// dropping tower/CTAF), bug #2 (remarks TAB_NAME), bug #3 tower attach.
// -----------------------------------------------------------------------
(function checkKDTO() {
  const rec = DATA.KDTO;
  ok("KDTO exists in dataset", !!rec);
  if (!rec) return;

  ok("KDTO is towered (tower !== null)", rec.tower !== null);
  eq("KDTO tower", rec.tower, {
    call: "DENTON", hours: "0600-2200", apchProvider: "D10", depProvider: "D10",
  });

  const rw1 = findRunway(rec, "18L/36R");
  ok("KDTO has runway 18L/36R", !!rw1);
  if (rw1) {
    eq("KDTO 18L/36R length", rw1.length, 7002);
    eq("KDTO 18L/36R width", rw1.width, 150);
    eq("KDTO 18L/36R gross wt SW/DW", [rw1.grossWtSW, rw1.grossWtDW], ["70", "100"]);
  }

  const rw2 = findRunway(rec, "18R/36L");
  ok("KDTO has runway 18R/36L", !!rw2);
  if (rw2) {
    eq("KDTO 18R/36L length", rw2.length, 5003);
    eq("KDTO 18R/36L width", rw2.width, 75);
  }

  // Bug #1: tower freq 119.95 must survive -- it used to be pushed past
  // the "first 20" cap by ~30 TRACON procedure rows that precede it in FRQ.csv.
  const twrFreq = rec.frequencies.find((f) => f.freq === "119.95" && f.use === "LCL/P");
  ok("KDTO tower frequency 119.95 (LCL/P) is present", !!twrFreq);
  const ctafFreq = rec.frequencies.find((f) => f.freq === "119.95" && f.use === "CTAF");
  ok("KDTO CTAF frequency 119.95 is present", !!ctafFreq);
  const atisFreq = rec.frequencies.find((f) => f.use === "ATIS");
  ok("KDTO ATIS frequency is present", !!atisFreq);

  // Bug #2: remarks must not be empty, and must match the AIRPORT-tab rows only.
  eq("KDTO remarks (exactly the 5 AIRPORT-tab rows, in file order)", rec.remarks, [
    "GENERAL_REMARK: RWY 18L DESIGNATED AS A CALM WIND RWY.",
    "GENERAL_REMARK: ARPT CLSD TO ULTRALIGHTS AND GLIDERS.",
    "GENERAL_REMARK: MOWING OPNS ON ARPT MAY-SEP.",
    "GENERAL_REMARK: FOR CD WHEN ATCT IS CLSD CTC LONE STAR APCH AT 972-615-2799.",
    "LGT_SKED: DUSK-DAWN WHEN ATCT CLSD, MIRL RWY 18L/36R & 18R/36L PRESET TO LOW INTST; TO INCR INTST & ACTVT MALSR RWY 18L - CTAF.",
  ]);
})();

// -----------------------------------------------------------------------
// KGAD -- Northeast Alabama Rgnl, AL. ATC_BASE has a NON-ATCT row for this
// field; bug #3 attached a tower object to it anyway. Must be null.
// -----------------------------------------------------------------------
(function checkKGAD() {
  const rec = DATA.KGAD;
  ok("KGAD exists in dataset", !!rec);
  if (!rec) return;
  eq("KGAD tower is null (ATC_BASE row is NON-ATCT)", rec.tower, null);
  eq("KGAD towerTypeCode raw passthrough", rec.towerTypeCode, "NON-ATCT");
  eq("KGAD attendance (2 segments)", rec.attendance, [
    { seq: "2", month: "ALL", days: "MON-FRI", hours: "0800-1800" },
    { seq: "3", month: "ALL", days: "SAT", hours: "0900-1700" },
  ]);
})();

// -----------------------------------------------------------------------
// KHYR -- Sawyer County, WI. Short field: 1088 ft turf crosswind runway.
// -----------------------------------------------------------------------
(function checkKHYR() {
  const rec = DATA.KHYR;
  ok("KHYR exists in dataset", !!rec);
  if (!rec) return;
  eq("KHYR tower is null (uncontrolled)", rec.tower, null);

  const rw = findRunway(rec, "16/34");
  ok("KHYR has short runway 16/34", !!rw);
  if (rw) {
    eq("KHYR 16/34 length", rw.length, 1088);
    eq("KHYR 16/34 width", rw.width, 120);
    eq("KHYR 16/34 surface code", rw.surface, "TURF");
    eq("KHYR 16/34 surface decoded", rw.surfaceDesc, "Turf/grass");
    eq("KHYR 16/34 has no edge lighting code (blank in source)", rw.lighting, "");
  }
})();

// -----------------------------------------------------------------------
// KALX -- Thomas C Russell Fld, AL. Runway 18 has a genuine 623 ft
// displaced threshold.
// -----------------------------------------------------------------------
(function checkKALX() {
  const rec = DATA.KALX;
  ok("KALX exists in dataset", !!rec);
  if (!rec) return;
  eq("KALX tower is null (uncontrolled)", rec.tower, null);

  const rw = findRunway(rec, "18/36");
  ok("KALX has runway 18/36", !!rw);
  if (rw) {
    const end18 = findEnd(rw, "18");
    ok("KALX runway 18 end exists", !!end18);
    if (end18) {
      eq("KALX end 18 displaced threshold length", end18.displacedThrLen, "623");
      eq("KALX end 18 hasDisplacedThr flag", end18.hasDisplacedThr, true);
    }
    const end36 = findEnd(rw, "36");
    ok("KALX runway 36 end exists", !!end36);
    if (end36) {
      eq("KALX end 36 has no displaced threshold", end36.hasDisplacedThr, false);
    }
  }
})();

// -----------------------------------------------------------------------
// KADS -- Addison, TX. Towered, busy reliever, but NOT FAR Part 139
// certificated -- no ARFF. Confirms ARFF absence is independent of tower
// status (a towered field can still have no ARFF).
// -----------------------------------------------------------------------
(function checkKADS() {
  const rec = DATA.KADS;
  ok("KADS exists in dataset", !!rec);
  if (!rec) return;
  ok("KADS is towered", rec.tower !== null);
  eq("KADS has no FAR 139 type (no ARFF)", rec.far139Type, "");
  eq("KADS has no ARFF cert type date", rec.arffCertTypeDate, "");
})();

console.log("\n" + checks + " checks, " + failures + " failure(s).");
if (failures > 0) {
  console.error("GOLDEN FIXTURE VERIFICATION FAILED.");
  process.exit(1);
}
console.log("Golden fixture verification OK.");
