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
//    (all thresholds null) and check the honesty guarantees.
// -----------------------------------------------------------------------
const report = ReportData.buildReport("KDTO", airportData, criteriaData, fboData);
ok("KDTO report built without error", !report.error);

ok("report.verified is false (criteria.json ships unreviewed)", report.verified === false);

const twrFreq = report.communications.find((c) => c.freq === "119.95" && c.use === "LCL/P");
ok("KDTO communications includes tower 119.95 (LCL/P)", !!twrFreq);
const ctafFreq = report.communications.find((c) => c.freq === "119.95" && c.use === "CTAF");
ok("KDTO communications includes CTAF 119.95", !!ctafFreq);

const rw1 = report.runwayFacts.find((r) => r.id === "18L/36R");
const rw2 = report.runwayFacts.find((r) => r.id === "18R/36L");
ok("KDTO runway 18L/36R present", !!rw1);
ok("KDTO runway 18R/36L present", !!rw2);
if (rw1) eq("KDTO 18L/36R length/width", [rw1.length, rw1.width], [7002, 150]);
if (rw2) eq("KDTO 18R/36L length/width", [rw2.length, rw2.width], [5003, 75]);

eq("KDTO facility.remarks has exactly 5 entries", report.facility.remarks.length, 5);

// Every runway-length headline verdict must be NOT EVALUATED (criteria.json
// has no threshold), which means every runway is a "qualifying" candidate
// and therefore gets a full criteria table -- this is the behavior that
// proves the tool refuses to silently exclude a runway it hasn't measured
// against a real number.
let allRows = [];
report.determinations.forEach((a) => {
  a.headline.forEach((h) => {
    ok(a.aircraftName + " " + h.runwayId + " headline is NOT EVALUATED (no threshold set)", h.result === "NOT EVALUATED");
  });
  a.criteriaTables.forEach((t) => { allRows = allRows.concat(t.rows); });
});
ok("at least one criteria table was produced per aircraft", report.determinations.every((a) => a.criteriaTables.length > 0));
ok("every criterion row reads NOT EVALUATED with criteria.json all-null (" + allRows.length + " rows checked)",
  allRows.length > 0 && allRows.every((r) => r.result === "NOT EVALUATED"));

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

  function xmlOk(desc, needle) {
    ok(desc, xml.indexOf(needle) !== -1);
  }
  xmlOk("docx XML contains tower frequency 119.95", "119.95");
  xmlOk("docx XML contains CRITERIA NOT VERIFIED banner", "CRITERIA NOT VERIFIED");
  xmlOk("docx XML contains NOT EVALUATED", "NOT EVALUATED");
  xmlOk("docx XML contains all 5 remarks (spot check calm wind remark)", "CALM WIND");
  xmlOk("docx XML contains runway 18L/36R", "18L/36R");
  xmlOk("docx XML contains runway 18R/36L", "18R/36L");

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
