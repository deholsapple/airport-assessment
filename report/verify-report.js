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
//    path. Everything else defaults conservatively.
// -----------------------------------------------------------------------
const reportDefault = ReportData.buildReport("KDTO", airportData, criteriaData, fboData, {});
ok("KDTO report (default extras) built without error", !reportDefault.error);
ok("report.verified is false (criteria.json ships unreviewed weight threshold)", reportDefault.verified === false);

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

// Both runways are paved -> hard-surface screen passes for both aircraft;
// weight-bearing threshold is null -> NOT EVALUATED; nothing else flagged
// -> should screen OK, not a performance determination.
reportDefault.determinations.forEach((d) => {
  ok(d.aircraftName + " bucket is SCREENS OK (paved runways, ILS present, towered)", d.bucket === ReportData.BUCKET.OK);
  eq(d.aircraftName + " has no reasons (clean screen)", d.reasons, []);
});
ok("no CLEAR DISQUALIFIER / SUITABLE / NOT SUITABLE language leaks into buckets", reportDefault.determinations.every((d) => ["CLEAR DISQUALIFIER", "REVIEW WITH CHIEF PILOT", "SCREENS OK — PERFORMANCE CALC REQUIRED"].indexOf(d.bucket) !== -1));

const surfaceRow18L = reportDefault.determinations[0].runwayScreens.find((r) => r.runwayId === "18L/36R").rows.find((r) => r.id === "surface");
eq("18L/36R surface screen: Asphalt, PASS (hard surface)", [surfaceRow18L.actualDisplay, surfaceRow18L.result], ["Asphalt", "PASS"]);
const weightRow18L = reportDefault.determinations[0].runwayScreens.find((r) => r.runwayId === "18L/36R").rows.find((r) => r.id === "wtBearingDW");
eq("18L/36R weight-bearing screen: NOT EVALUATED (threshold still null)", weightRow18L.result, "NOT EVALUATED");

// Declared distances are displayed facts now, not a gate -- confirm they
// still appear correctly (18L LDA 6,502 ft vs 7,002 ft physical length).
const end18L = reportDefault.endFacts.find((e) => e.endId === "18L");
eq("KDTO 18L declared distances still surfaced as facts", [end18L.tora, end18L.toda, end18L.asda, end18L.lda], ["7002", "7002", "6502", "6502"]);
const end18R = reportDefault.endFacts.find((e) => e.endId === "18R");
eq("KDTO 18R (no declared distances) reads null, not a length fallback", [end18R.tora, end18R.toda, end18R.asda, end18R.lda], [null, null, null, null]);

// -----------------------------------------------------------------------
// 2. Communications: operational set only (Fix 1 carried forward).
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
// 3. Pilot-answered extras: recency=no, winter-night=yes at an uncontrolled
//    field with no on-field weather (KGAD has an AWOS, so use KHYR? KHYR
//    has an ASOS too. Use a scenario where we force remoteUnmonitored via
//    KALX, which is uncontrolled -- check whether it has on-field weather.)
// -----------------------------------------------------------------------
console.log("\n--- Pilot-answered extras & winter-night hard flag ---");
// KAIV (George Downer, AL) is uncontrolled with no on-field ASOS/AWOS --
// a real, verified instance of the §6.1.6.7 remote/unmonitored condition
// (KDTO/KGAD/KHYR/KALX/KADS all happen to have on-field weather reporting,
// so this scenario needs a different real airport to exercise on real data
// rather than only via synthetic overrides).
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
// 4. FBO editable + lookup link, PIC comments + sign-off (carried forward).
// -----------------------------------------------------------------------
console.log("\n--- FBO / PIC comments / sign-off ---");
const extras = {
  manualFbo: [{ name: "Sheltair DTO", phone: "940-435-2621", fuel: "Jet A, 100LL", services: "Hangar, crew car, catering" }],
  picComments: "Verified runway 18L/36R suitable for planned ops. No NOTAMs affecting arrival.",
  picName: "J. Smith",
  chiefPilotName: "R. Jones",
};
const report = ReportData.buildReport("KDTO", airportData, criteriaData, fboData, extras);
eq("KDTO fboLookupUrl resolves to KDTO's AirNav page", report.fboLookupUrl, "https://www.airnav.com/airport/KDTO");
ok("manual FBO entry (Sheltair DTO) present with typed values, not CONFIRM", report.fbo.length === 1 && report.fbo[0].phone === "940-435-2621");
eq("picComments carried through to report model", report.picComments, extras.picComments);
eq("picName / chiefPilotName carried through", [report.picName, report.chiefPilotName], [extras.picName, extras.chiefPilotName]);

console.log("\n" + checks + " model checks, " + failures + " failure(s).");

// -----------------------------------------------------------------------
// 5. Build the actual .docx (same builder the browser uses) and sanity
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
  const rels = fs.readFileSync(path.join(extractDir, "word", "_rels", "document.xml.rels"), "utf8");

  function xmlOk(desc, needle) { ok(desc, xml.indexOf(needle) !== -1); }
  xmlOk("docx XML leads with FIRST-PASS SCREENING ONLY banner", "FIRST-PASS SCREENING ONLY");
  xmlOk("docx XML mentions ForeFlight performance calculation requirement", "Performance calculation (ForeFlight) required");
  xmlOk("docx XML mentions Chief Pilot review of marginal items", "reviewed with the Chief Pilot");
  xmlOk("docx XML title is AIRPORT SCREENING REPORT (not SUITABILITY ASSESSMENT)", "AIRPORT SCREENING REPORT");
  ok("docx XML does not contain the old SUITABLE/NOT SUITABLE determination language", xml.indexOf(">SUITABLE<") === -1 && xml.indexOf("NOT SUITABLE") === -1);
  xmlOk("docx XML contains SCREENS OK bucket text", "SCREENS OK");
  xmlOk("docx XML contains tower frequency 119.95", "119.95");
  xmlOk("docx XML contains declared distance 6,502 (18L LDA, displayed as fact)", "6,502");
  xmlOk("docx XML contains all 5 remarks (spot check calm wind remark)", "CALM WIND");
  ok("docx XML does not contain a named procedure tag (GREGS STAR)", xml.indexOf("GREGS") === -1);
  xmlOk("docx XML contains alternate minimums 600-2 with FOM citation", "600-2");
  xmlOk("docx XML contains FOM §4.1.3.3 citation", "4.1.3.3");
  xmlOk("docx XML contains night circling FOM §6.1.6.2 citation", "6.1.6.2");
  xmlOk("docx XML contains night-capability / remote-ops FOM §6.1.6.7 citation", "6.1.6.7");
  xmlOk("docx XML contains 2-year recency ARA note (§2.7.4)", "2.7.4");
  xmlOk("docx XML contains the manual FBO name (Sheltair DTO)", "Sheltair DTO");
  xmlOk("docx XML contains the AirNav lookup URL", "airnav.com/airport/KDTO");
  ok("docx relationships include the AirNav hyperlink", rels.indexOf("airnav.com/airport/KDTO") !== -1);
  xmlOk("docx XML contains PIC comments text", "Verified runway 18L/36R suitable");
  xmlOk("docx XML contains PIC name in sign-off", "J. Smith");
  xmlOk("docx XML contains Chief Pilot name in sign-off", "R. Jones");
  xmlOk("docx XML contains SIGN-OFF section", "SIGN-OFF");
  xmlOk("docx XML contains PIC NARRATIVE section", "PIC NARRATIVE");
  xmlOk("docx XML contains NIGHT OPERATIONS & RECENCY section", "NIGHT OPERATIONS");

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
