// End-to-end verification for the report generator: builds a real KDTO
// .docx (same code path the browser uses) and checks both the computed
// report model and the actual document XML for the facts that must be
// right every time.
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
// 1. Build the report model for KDTO with criteria.json exactly as shipped
//    (all thresholds null), a manual FBO entry, PIC comments and sign-off
//    names -- and check the honesty guarantees plus the five real-world
//    fixes.
// -----------------------------------------------------------------------
const extras = {
  manualFbo: [{ name: "Sheltair DTO", phone: "940-435-2621", fuel: "Jet A, 100LL", services: "Hangar, crew car, catering" }],
  picComments: "Verified runway 18L/36R suitable for planned ops. No NOTAMs affecting arrival.",
  picName: "J. Smith",
  chiefPilotName: "R. Jones",
};
const report = ReportData.buildReport("KDTO", airportData, criteriaData, fboData, extras);
ok("KDTO report built without error", !report.error);
ok("report.verified is false (criteria.json ships unreviewed)", report.verified === false);

// --- Fix 1: communications filtered to the operational set -------------
console.log("\nKDTO communications rows (" + report.communications.length + "):");
report.communications.forEach((c) => console.log("  " + c.freq + " " + c.use + " (" + c.facilityType + ")"));
ok("KDTO communications is roughly 7 rows, not ~38 (" + report.communications.length + ")", report.communications.length >= 6 && report.communications.length <= 9);
ok("no STAR/DP/RNAV named-procedure rows survive", report.communications.every((c) => !/\b(STAR|DP|RNAV)\b/.test(c.use.toUpperCase())));
const twrFreq = report.communications.find((c) => c.freq === "119.95" && c.use === "LCL/P");
ok("KDTO communications includes tower 119.95 (LCL/P)", !!twrFreq);
const ctafFreq = report.communications.find((c) => c.freq === "119.95" && c.use === "CTAF");
ok("KDTO communications includes CTAF 119.95", !!ctafFreq);
const dupCheck = new Set(report.communications.map((c) => c.freq + "|" + c.use));
ok("no duplicate freq+use rows", dupCheck.size === report.communications.length);

eq("KDTO facility.remarks has exactly 5 entries", report.facility.remarks.length, 5);

// --- Fix 2 & 3: merged runway/criteria evaluation using declared distances
const end18L = report.endFacts.find((e) => e.endId === "18L");
const end36R = report.endFacts.find((e) => e.endId === "36R");
const end18R = report.endFacts.find((e) => e.endId === "18R");
ok("KDTO 18L end present in endFacts", !!end18L);
ok("KDTO 36R end present in endFacts", !!end36R);
ok("KDTO 18R end present in endFacts", !!end18R);
if (end18L) {
  eq("KDTO 18L declared distances", [end18L.tora, end18L.toda, end18L.asda, end18L.lda], ["7002", "7002", "6502", "6502"]);
  ok("KDTO 18L LDA (6,502 ft) differs from its physical runway length (7,002 ft)", end18L.lda !== "7002");
}
if (end18R) {
  eq("KDTO 18R has no declared distances published (NOT AVAILABLE, no fallback to length)", [end18R.tora, end18R.toda, end18R.asda, end18R.lda], [null, null, null, null]);
}

let allRows = [];
report.determinations.forEach((a) => {
  a.criteriaTables.forEach((t) => { allRows = allRows.concat(t.rows); });
});
ok("at least one criteria table was produced per aircraft", report.determinations.every((a) => a.criteriaTables.length > 0));
ok("every criterion row reads NOT EVALUATED with criteria.json all-null (" + allRows.length + " rows checked)",
  allRows.length > 0 && allRows.every((r) => r.result === "NOT EVALUATED"));
// 18R/36L has no declared distances -- its takeoff/landing distance rows
// must still read NOT EVALUATED (threshold is null), never UNKNOWN, since
// threshold-missing always wins regardless of data availability.
const end18RTable = report.determinations[0].criteriaTables.find((t) => t.endId === "18R");
ok("18R criteria table produced despite missing declared distances", !!end18RTable);
if (end18RTable) {
  const takeoffRow = end18RTable.rows.find((r) => r.id === "takeoffDistance");
  ok("18R takeoff distance row is NOT EVALUATED (threshold null wins over missing data)", takeoffRow.result === "NOT EVALUATED");
}

// --- Fix 4: FBO editable + lookup link ----------------------------------
eq("KDTO fboLookupUrl resolves to KDTO's AirNav page", report.fboLookupUrl, "https://www.airnav.com/airport/KDTO");
ok("manual FBO entry (Sheltair DTO) present with typed values, not CONFIRM", report.fbo.length === 1 && report.fbo[0].phone === "940-435-2621");
const noManual = ReportData.buildReport("KDTO", airportData, criteriaData, fboData, {});
ok("without manual FBO entries, fbo-data.js (currently empty for KDTO) is used and yields no rows", noManual.fbo.length === 0);

// --- Fix 5: PIC comments + sign-off --------------------------------------
eq("picComments carried through to report model", report.picComments, extras.picComments);
eq("picName / chiefPilotName carried through", [report.picName, report.chiefPilotName], [extras.picName, extras.chiefPilotName]);

console.log("\n" + checks + " model checks, " + failures + " failure(s).");

// -----------------------------------------------------------------------
// 2. Build the actual .docx (same builder the browser uses) and sanity
//    check the real document XML, not just the intermediate JS model.
// -----------------------------------------------------------------------
async function buildDocxAndVerify() {
  const doc = DocxBuilder.build(report, docx);
  const outDir = path.join(__dirname, "out");
  fs.mkdirSync(outDir, { recursive: true });
  const outFile = path.join(outDir, "KDTO-suitability-verify.docx");
  const buf = await docx.Packer.toBuffer(doc);
  fs.writeFileSync(outFile, buf);
  console.log("\nWrote " + outFile + " (" + buf.length + " bytes)");

  const extractDir = path.join(outDir, "extracted");
  fs.rmSync(extractDir, { recursive: true, force: true });
  fs.mkdirSync(extractDir, { recursive: true });
  execSync(`unzip -o -q "${outFile}" -d "${extractDir}"`);
  const xml = fs.readFileSync(path.join(extractDir, "word", "document.xml"), "utf8");
  const rels = fs.readFileSync(path.join(extractDir, "word", "_rels", "document.xml.rels"), "utf8");

  function xmlOk(desc, needle) {
    ok(desc, xml.indexOf(needle) !== -1);
  }
  xmlOk("docx XML contains tower frequency 119.95", "119.95");
  xmlOk("docx XML contains CRITERIA NOT VERIFIED banner", "CRITERIA NOT VERIFIED");
  xmlOk("docx XML contains NOT EVALUATED", "NOT EVALUATED");
  xmlOk("docx XML contains all 5 remarks (spot check calm wind remark)", "CALM WIND");
  xmlOk("docx XML contains declared distance 6,502 (18L LDA)", "6,502");
  ok("docx XML does not contain a named procedure tag (GREGS STAR)", xml.indexOf("GREGS") === -1);
  xmlOk("docx XML contains the manual FBO name (Sheltair DTO)", "Sheltair DTO");
  xmlOk("docx XML contains the AirNav lookup URL", "airnav.com/airport/KDTO");
  ok("docx relationships include the AirNav hyperlink", rels.indexOf("airnav.com/airport/KDTO") !== -1);
  xmlOk("docx XML contains PIC comments text", "Verified runway 18L/36R suitable");
  xmlOk("docx XML contains PIC name in sign-off", "J. Smith");
  xmlOk("docx XML contains Chief Pilot name in sign-off", "R. Jones");
  xmlOk("docx XML contains SIGN-OFF section", "SIGN-OFF");
  xmlOk("docx XML contains PIC NARRATIVE section", "PIC NARRATIVE");

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
