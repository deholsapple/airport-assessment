// End-to-end verification for the airport screening report: builds a real
// KDTO .docx (same code path the browser uses) and checks both the
// computed report model and the actual document XML for the facts that
// must be right every time.
//
// Usage: node report/verify-report.js

const fs = require("fs");
const path = require("path");
const vm = require("vm");
const { execSync } = require("child_process");

const docx = require("docx");
const ReportData = require("./report-data.js");
const DocxBuilder = require("./docx-builder.js");

const ROOT = path.join(__dirname, "..");

function loadWindowGlobal(file, globalName) {
  const src = fs.readFileSync(path.join(ROOT, file), "utf8");
  const sandbox = { window: {} };
  vm.createContext(sandbox);
  vm.runInContext(src, sandbox, { filename: file });
  return sandbox.window[globalName];
}

const airportData = loadWindowGlobal("airport-data.js", "NASR_AIRPORT_DATA");
const fboData = loadWindowGlobal("fbo-data.js", "FBO_DATA");
const criteriaData = JSON.parse(fs.readFileSync(path.join(ROOT, "criteria.json"), "utf8"));

let failures = 0;
let checks = 0;
function ok(desc, cond) {
  checks++;
  if (!cond) { failures++; console.error("FAIL: " + desc); } else { console.log("ok - " + desc); }
}
function eq(desc, actual, expected) {
  checks++;
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a !== e) { failures++; console.error("FAIL: " + desc + "\n  expected: " + e + "\n  actual:   " + a); } else { console.log("ok - " + desc); }
}

// -----------------------------------------------------------------------
// 1. Default state: no pilot answers given (all default to "unknown").
//    KDTO has an ILS on 18L, so instrument approach should auto-derive to
//    "Precision or APV" even with no override -- this is the auto-derive
//    path. Both runways are well above both airframes' weight-bearing
//    floor (D-100 = 100,000 lb vs. 24,286 / 26,689 lb).
// -----------------------------------------------------------------------
const reportDefault = ReportData.buildReport("KDTO", airportData, criteriaData, fboData, {});
ok("KDTO report (default extras) built without error", !reportDefault.error);
ok("report.verified is true (both surface and weight-bearing thresholds are now real numbers)", reportDefault.verified === true);

console.log("\n--- Default-state checks (no pilot answers) ---");
eq("KDTO instrument approach auto-derives to precisionOrApv (18L has ILS)", reportDefault.nightOps.instrumentApproach.auto, "precisionOrApv");
eq("KDTO effective approach type = auto-derived (no override)", reportDefault.nightOps.instrumentApproach.effective, "precisionOrApv");
ok("KDTO instrument approach NOT flagged as manual-entry-required (auto-derived)", reportDefault.nightOps.instrumentApproach.manualEntryRequired === false);
eq("KDTO alternate minimums = 600-2 (precision/APV, FOM §4.1.3.3)", reportDefault.nightOps.alternateMinimums, { ceilingFt: 600, visibilityMi: 2, citation: "FOM §4.1.3.3" });
ok("KDTO is towered so NOT remote/unmonitored", reportDefault.nightOps.remoteUnmonitored === false);
eq("KDTO 2-year recency defaults to unknown", reportDefault.nightOps.recency.answer, "unknown");
eq("KDTO winter-night-ops defaults to unknown, no hard flag", [reportDefault.nightOps.winterNightOps, reportDefault.nightOps.winterNightHardFlag], ["unknown", false]);
ok("KDTO wind indicator is lighted (night capability)", reportDefault.nightOps.nightCapability.windIndicatorLighted === true);
ok("KDTO has runway edge lighting (night capability)", reportDefault.nightOps.nightCapability.edgeLightingPresent === true);
ok("night circling note present and cites FOM §6.1.6.2", reportDefault.nightOps.nightCirclingNote.indexOf("§6.1.6.2") !== -1);

// Both runways are paved and both clear the weight-bearing floor -> clean
// screen, no reasons, SCREENS OK for both airframes.
reportDefault.determinations.forEach((d) => {
  ok(d.aircraftName + " bucket is SCREENS OK (paved runways, weight-bearing clears floor, ILS present, towered)", d.bucket === ReportData.BUCKET.OK);
  eq(d.aircraftName + " has no reasons (clean screen)", d.reasons, []);
});
ok("no CLEAR DISQUALIFIER / SUITABLE / NOT SUITABLE language leaks into buckets", reportDefault.determinations.every((d) => ["CLEAR DISQUALIFIER", "REVIEW WITH CHIEF PILOT", "SCREENS OK — PERFORMANCE CALC REQUIRED"].indexOf(d.bucket) !== -1));

const rs18L_falcon = reportDefault.determinations.find((d) => d.aircraftId === "falcon50ex").runwayScreens.find((r) => r.runwayId === "18L/36R");
const surfaceRow18L = rs18L_falcon.rows.find((r) => r.id === "surface");
eq("18L/36R surface screen: Asphalt, PASS (hard surface)", [surfaceRow18L.actualDisplay, surfaceRow18L.result], ["Asphalt", "PASS"]);
const weightRow18L_falcon = rs18L_falcon.rows.find((r) => r.id === "wtBearingDW");
eq("18L/36R weight-bearing screen (Falcon 50EX): 100,000 lb clears 24,286 lb floor -> PASS", [weightRow18L_falcon.actualDisplay, weightRow18L_falcon.result], ["100,000 lb", "PASS"]);
const rs18L_g280 = reportDefault.determinations.find((d) => d.aircraftId === "g280").runwayScreens.find((r) => r.runwayId === "18L/36R");
const weightRow18L_g280 = rs18L_g280.rows.find((r) => r.id === "wtBearingDW");
eq("18L/36R weight-bearing screen (G280): 100,000 lb clears 26,689 lb floor -> PASS", [weightRow18L_g280.actualDisplay, weightRow18L_g280.result], ["100,000 lb", "PASS"]);

// Runway distances (physical length, declared distances) are displayed as
// reference facts regardless of screen result, so a pilot can judge the
// heavier-weight/crosswind-runway case themselves.
const rw18L = reportDefault.runwayFacts.find((r) => r.id === "18L/36R");
eq("KDTO 18L/36R physical length displayed as a fact", rw18L.length, 7002);
const end18L = reportDefault.endFacts.find((e) => e.endId === "18L");
eq("KDTO 18L declared distances still surfaced as facts", [end18L.tora, end18L.toda, end18L.asda, end18L.lda], ["7002", "7002", "6502", "6502"]);
const end18R = reportDefault.endFacts.find((e) => e.endId === "18R");
eq("KDTO 18R (no declared distances) reads null, not a length fallback", [end18R.tora, end18R.toda, end18R.asda, end18R.lda], [null, null, null, null]);

// -----------------------------------------------------------------------
// 2. Weight-bearing disqualifier -- real, verified airport, not synthetic.
//    KITR (Kit Carson County, CO) is a single-runway field with a
//    published DW rating of 17 (17,000 lb), well below both airframes'
//    lightest landing weight (24,286 / 26,689 lb). Its surface IS paved
//    (Asphalt), which isolates the weight-bearing gate specifically --
//    proof the disqualifier is driven by weight, not surface.
// -----------------------------------------------------------------------
console.log("\n--- Weight-bearing disqualifier (real airport, not synthetic) ---");
const kitr = ReportData.buildReport("KITR", airportData, criteriaData, fboData, {});
ok("KITR report built without error", !kitr.error);
const kitrRunway = kitr.runwayFacts.find((r) => r.id === "15/33");
ok("KITR has a single paved runway 15/33", !!kitrRunway && kitrRunway.surface === "Asphalt");
const kitrScreen = kitr.determinations[0].runwayScreens.find((r) => r.runwayId === "15/33");
const kitrSurface = kitrScreen.rows.find((r) => r.id === "surface");
const kitrWeight = kitrScreen.rows.find((r) => r.id === "wtBearingDW");
eq("KITR 15/33 surface screen PASSES (it IS paved -- isolates the weight gate)", kitrSurface.result, "PASS");
eq("KITR 15/33 weight-bearing actual: 17,000 lb (raw NASR DW=17, thousands)", kitrWeight.actualDisplay, "17,000 lb");
eq("KITR 15/33 weight-bearing screen FAILS for both airframes (17,000 lb < both floors)", kitrScreen.verdict, "FAIL");
kitr.determinations.forEach((d) => {
  ok(d.aircraftName + " at KITR is CLEAR DISQUALIFIER (only runway fails weight-bearing floor)", d.bucket === ReportData.BUCKET.DISQUALIFIER);
});
ok("KITR disqualifying findings cite the weight-bearing screen", kitr.disqualifyingFindings.some((f) => /hard \(paved\) surface \/ meets the weight-bearing screen/.test(f)));

// -----------------------------------------------------------------------
// 3. Blank NASR weight-bearing data must read NOT AVAILABLE / UNKNOWN --
//    never silently pass (treated as clearing the floor) and never
//    silently fail (treated as zero). KADS's runway has no blank DW in
//    our fixture set, so use a real airport whose runway has a blank
//    GROSS_WT_DW to exercise this path on real data.
// -----------------------------------------------------------------------
console.log("\n--- Blank weight-bearing data (real airport) ---");
let blankDwIcao = null, blankDwRunwayId = null;
for (const icao of Object.keys(airportData)) {
  const rec = airportData[icao];
  const rw = (rec.runways || []).find((r) => r.surface && r.surface.indexOf("ASPH") !== -1 && !r.grossWtDW);
  if (rw) { blankDwIcao = icao; blankDwRunwayId = rw.id; break; }
}
ok("found a real paved-runway airport with blank NASR weight-bearing data", !!blankDwIcao);
if (blankDwIcao) {
  const blankReport = ReportData.buildReport(blankDwIcao, airportData, criteriaData, fboData, {});
  const blankScreen = blankReport.determinations[0].runwayScreens.find((r) => r.runwayId === blankDwRunwayId);
  const blankWeight = blankScreen.rows.find((r) => r.id === "wtBearingDW");
  ok(blankDwIcao + " " + blankDwRunwayId + " weight-bearing reads UNKNOWN, not PASS or FAIL (blank != 0, blank != clears-floor)", blankWeight.result === "UNKNOWN");
  eq(blankDwIcao + " " + blankDwRunwayId + " weight-bearing actual is NOT AVAILABLE", blankWeight.actualDisplay, "NOT AVAILABLE");
  ok(blankDwIcao + " logs a weight-bearing verification item", blankReport.verificationItems.some((i) => /dual-wheel weight bearing not in NASR/.test(i.item)));
  ok(blankDwIcao + " bucket is not silently SCREENS OK (unknown weight-bearing routes to review)", blankReport.determinations.every((d) => d.bucket !== ReportData.BUCKET.OK) || blankScreen.verdict !== "UNKNOWN");
}

// -----------------------------------------------------------------------
// 4. Communications: operational set only (carried forward, unchanged).
// -----------------------------------------------------------------------
console.log("\n--- Communications ---");
console.log("KDTO communications rows (" + reportDefault.communications.length + "):");
reportDefault.communications.forEach((c) => console.log("  " + c.freq + " " + c.use + " (" + c.facilityType + ")"));
ok("KDTO communications is roughly 7 rows, not ~38 (" + reportDefault.communications.length + ")", reportDefault.communications.length >= 6 && reportDefault.communications.length <= 9);
ok("no STAR/DP/RNAV named-procedure rows survive", reportDefault.communications.every((c) => !/\b(STAR|DP|RNAV)\b/.test(c.use.toUpperCase())));
ok("KDTO communications includes tower 119.95 (LCL/P)", !!reportDefault.communications.find((c) => c.freq === "119.95" && c.use === "LCL/P"));
ok("KDTO communications includes CTAF 119.95", !!reportDefault.communications.find((c) => c.freq === "119.95" && c.use === "CTAF"));

eq("KDTO facility.remarks has exactly 5 entries", reportDefault.facility.remarks.length, 5);

// -----------------------------------------------------------------------
// 5. Pilot-answered extras: remote/unmonitored winter-night hard flag,
//    2-year recency, instrument approach override (carried forward).
// -----------------------------------------------------------------------
console.log("\n--- Pilot-answered extras & winter-night hard flag ---");
// KAIV (George Downer, AL) is uncontrolled with no on-field ASOS/AWOS --
// a real, verified instance of the §6.1.6.7 remote/unmonitored condition.
const kaiv = ReportData.buildReport("KAIV", airportData, criteriaData, fboData, { winterNightOps: "yes" });
ok("KAIV is uncontrolled with no on-field weather reporting (real NASR data)", kaiv.nightOps.remoteUnmonitored === true);
ok("KAIV + winterNightOps=yes + remote/unmonitored -> CLEAR DISQUALIFIER hard flag", kaiv.nightOps.winterNightHardFlag === true);
ok("KAIV determinations show CLEAR DISQUALIFIER when hard flag set", kaiv.determinations.every((d) => d.bucket === ReportData.BUCKET.DISQUALIFIER));
ok("KAIV disqualifying findings mention winter-season night operation", kaiv.disqualifyingFindings.some((f) => /winter-season night operation/.test(f)));

const kaivNoWinter = ReportData.buildReport("KAIV", airportData, criteriaData, fboData, { winterNightOps: "no" });
ok("Same remote/unmonitored condition WITHOUT winterNightOps=yes does not hard-disqualify", kaivNoWinter.nightOps.winterNightHardFlag === false);
ok("...but still drives REVIEW WITH CHIEF PILOT (real risk, just not the hard-flag scenario)", kaivNoWinter.determinations.every((d) => d.bucket === ReportData.BUCKET.REVIEW));

// Recency = no -> informational note + verification item, does not gate bucket.
const kdtoRecencyNo = ReportData.buildReport("KDTO", airportData, criteriaData, fboData, { recentOperation: "no" });
ok("recentOperation=no produces an ARA-required note", /new Airport Risk Assessment/.test(kdtoRecencyNo.nightOps.recency.note));
ok("recentOperation=no logs a verification item", kdtoRecencyNo.verificationItems.some((i) => /2-year recency/.test(i.item)));
ok("recentOperation=no does not by itself change the screening bucket", kdtoRecencyNo.determinations.every((d) => d.bucket === ReportData.BUCKET.OK));

// Instrument approach override: force "none" (VFR only) and confirm the
// VFR-only flag + no alternate minimums + it becomes a review reason.
const kdtoVfrOnly = ReportData.buildReport("KDTO", airportData, criteriaData, fboData, { instrumentApproachOverride: "none" });
eq("Override to 'none' sets vfrOnlyFlag", kdtoVfrOnly.nightOps.vfrOnlyFlag, true);
eq("Override to 'none' clears alternate minimums", kdtoVfrOnly.nightOps.alternateMinimums, null);
ok("Override is recorded as overridden vs. the ILS-derived auto value", kdtoVfrOnly.nightOps.instrumentApproach.overridden === true);

// Unknown approach type (simulate an airport with no ground navaid at all)
// should log a verification item and drive REVIEW WITH CHIEF PILOT.
const kdtoUnknownApch = ReportData.buildReport("KDTO", airportData, criteriaData, fboData, { instrumentApproachOverride: "unknown" });
ok("Forced-unknown approach type logs a verification item", kdtoUnknownApch.verificationItems.some((i) => /Instrument approach availability not confirmed/.test(i.item)));
ok("Forced-unknown approach type drives REVIEW WITH CHIEF PILOT", kdtoUnknownApch.determinations.every((d) => d.bucket === ReportData.BUCKET.REVIEW));

// -----------------------------------------------------------------------
// 6. FBO stays editable-section-only (Section 6 display removed); PIC
//    comments + sign-off carried forward.
// -----------------------------------------------------------------------
console.log("\n--- FBO (no display section) / PIC comments / sign-off ---");
ok("report model has no fboLookupUrl (dead field removed with Section 6)", reportDefault.fboLookupUrl === undefined);
const extras = {
  manualFbo: [{ name: "Sheltair DTO", phone: "940-435-2621", fuel: "Jet A, 100LL", services: "Hangar, crew car, catering" }],
  picComments: "Verified runway 18L/36R suitable for planned ops. No NOTAMs affecting arrival.",
  picName: "J. Smith",
  chiefPilotName: "R. Jones",
};
const report = ReportData.buildReport("KDTO", airportData, criteriaData, fboData, extras);
ok("manual FBO entry (Sheltair DTO) present with typed values, not CONFIRM -- still tracked even though not displayed", report.fbo.length === 1 && report.fbo[0].phone === "940-435-2621");
eq("picComments carried through to report model", report.picComments, extras.picComments);
eq("picName / chiefPilotName carried through", [report.picName, report.chiefPilotName], [extras.picName, extras.chiefPilotName]);

console.log("\n" + checks + " model checks, " + failures + " failure(s).");

// -----------------------------------------------------------------------
// 7. Build the actual .docx (same builder the browser uses) and sanity
//    check the real document XML, not just the intermediate JS model.
// -----------------------------------------------------------------------
async function buildDocxAndVerify() {
  const doc = DocxBuilder.build(report, docx);
  const outDir = path.join(__dirname, "out");
  fs.mkdirSync(outDir, { recursive: true });
  const outFile = path.join(outDir, "KDTO-screening-verify.docx");
  const buf = await docx.Packer.toBuffer(doc);
  fs.writeFileSync(outFile, buf);
  console.log("\nWrote " + outFile + " (" + buf.length + " bytes)");

  const extractDir = path.join(outDir, "extracted");
  fs.rmSync(extractDir, { recursive: true, force: true });
  fs.mkdirSync(extractDir, { recursive: true });
  execSync(`unzip -o -q "${outFile}" -d "${extractDir}"`);
  const xml = fs.readFileSync(path.join(extractDir, "word", "document.xml"), "utf8");

  function xmlOk(desc, needle) { ok(desc, xml.indexOf(needle) !== -1); }
  xmlOk("docx XML leads with FIRST-PASS SCREENING ONLY banner", "FIRST-PASS SCREENING ONLY");
  xmlOk("docx XML mentions ForeFlight performance calculation requirement", "Performance calculation (ForeFlight) required");
  xmlOk("docx XML mentions Chief Pilot review of marginal items", "reviewed with the Chief Pilot");
  xmlOk("docx XML title is AIRPORT SCREENING REPORT (not SUITABILITY ASSESSMENT)", "AIRPORT SCREENING REPORT");
  ok("docx XML does not contain the old SUITABLE/NOT SUITABLE determination language", xml.indexOf(">SUITABLE<") === -1 && xml.indexOf("NOT SUITABLE") === -1);
  ok("docx XML has no CRITERIA NOT VERIFIED banner (criteria.json is now fully reviewed)", xml.indexOf("CRITERIA NOT VERIFIED") === -1);
  xmlOk("docx XML contains SCREENS OK bucket text", "SCREENS OK");
  xmlOk("docx XML contains tower frequency 119.95", "119.95");
  xmlOk("docx XML contains physical runway length 7,002 ft (displayed fact)", "7,002 ft long");
  xmlOk("docx XML contains declared distance 6,502 (18L LDA, displayed as fact)", "6,502");
  xmlOk("docx XML contains all 5 remarks (spot check calm wind remark)", "CALM WIND");
  xmlOk("docx XML contains weight-bearing actual 100,000 lb", "100,000 lb");
  xmlOk("docx XML contains weight-bearing floor 24,286 lb (Falcon 50EX)", "24,286");
  xmlOk("docx XML contains weight-bearing floor 26,689 lb (G280)", "26,689");
  ok("docx XML does not contain a named procedure tag (GREGS STAR)", xml.indexOf("GREGS") === -1);
  xmlOk("docx XML contains alternate minimums 600-2 with FOM citation", "600-2");
  xmlOk("docx XML contains FOM §4.1.3.3 citation", "4.1.3.3");
  xmlOk("docx XML contains night circling FOM §6.1.6.2 citation", "6.1.6.2");
  xmlOk("docx XML contains night-capability / remote-ops FOM §6.1.6.7 citation", "6.1.6.7");
  xmlOk("docx XML contains 2-year recency ARA note (§2.7.4)", "2.7.4");
  ok("docx XML has no SERVICES AND FBO section (removed as redundant)", xml.indexOf("SERVICES AND FBO") === -1);
  ok("docx XML has no FBO table content (Sheltair DTO must not appear in the printed report)", xml.indexOf("Sheltair") === -1);
  xmlOk("docx XML contains PIC comments text", "Verified runway 18L/36R suitable");
  xmlOk("docx XML contains PIC name in sign-off", "J. Smith");
  xmlOk("docx XML contains Chief Pilot name in sign-off", "R. Jones");
  xmlOk("docx XML contains SIGN-OFF section", "SIGN-OFF");
  xmlOk("docx XML contains PIC NARRATIVE section", "PIC NARRATIVE");
  xmlOk("docx XML contains NIGHT OPERATIONS & RECENCY section", "NIGHT OPERATIONS");
  // Section numbering: Runway Eval (4) -> Night Ops (5) -> Disqualifying (6)
  // -> Verification (7) -> PIC Narrative (8) -> Sign-off (9), no gap for FBO.
  xmlOk("docx XML numbers Disqualifying Findings as section 6", "6.  DISQUALIFYING FINDINGS");
  xmlOk("docx XML numbers Items Requiring Verification as section 7", "7.  ITEMS REQUIRING VERIFICATION");
  xmlOk("docx XML numbers PIC Narrative as section 8", "8.  PIC NARRATIVE");
  xmlOk("docx XML numbers Sign-off as section 9", "9.  SIGN-OFF");

  console.log("\n" + checks + " total checks, " + failures + " failure(s).");
  if (failures > 0) {
    console.error("REPORT VERIFICATION FAILED.");
    process.exit(1);
  }
  console.log("Report verification OK.");
}

buildDocxAndVerify().catch((err) => {
  console.error(err);
  process.exit(1);
});
