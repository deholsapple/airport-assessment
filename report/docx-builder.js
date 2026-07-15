// Renders a report-data.js report model into a docx.Document -- a first-
// pass screening report, not a performance-suitability determination.
// Takes the docx library as a parameter (window.docx in the browser,
// require("docx") under Node) so this file has no environment dependency.
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.DocxBuilder = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var NAVY = "1F3864";
  var BANNER_BG = "FFF2CC";
  var INFO_BANNER_BG = "DCE6F1";
  var RED = "C00000";
  var GOLD = "7F6000";
  var AMBER = "BF8F00";
  var GREEN = "375623";
  var GRAY_TEXT = "404040";
  var GRAY_NOTE = "595959";
  var WHITE = "FFFFFF";

  var CONTENT_WIDTH = 9360; // 6.5in at 1440 dxa/in
  var FONT = "Calibri";

  var RESULT_COLOR = {
    PASS: GREEN, CAUTION: AMBER, FAIL: RED, UNKNOWN: GOLD, "NOT EVALUATED": GRAY_NOTE,
  };
  var BUCKET_COLOR = {
    "CLEAR DISQUALIFIER": RED,
    "REVIEW WITH CHIEF PILOT": AMBER,
    "SCREENS OK — PERFORMANCE CALC REQUIRED": GREEN,
  };

  function build(report, D) {
    var children = [];

    children.push(banner(D,
      "FIRST-PASS SCREENING ONLY",
      "Performance calculation (ForeFlight) required before dispatch. Marginal items are to be reviewed with the Chief Pilot. " +
      "This report never replaces PIC judgment and does not substitute for a DOA/ForeFlight determination.",
      NAVY, INFO_BANNER_BG));

    if (!report.verified) {
      children.push(banner(D,
        "CRITERIA NOT VERIFIED",
        "Weight-bearing thresholds in criteria.json have not been reviewed against a real company minimum. Every criterion depending on an unset threshold reads NOT EVALUATED.",
        RED, BANNER_BG));
    }

    children.push(new D.Paragraph({
      children: [new D.TextRun({ text: "AIRPORT SCREENING REPORT", bold: true, color: NAVY, font: FONT, size: 32 })],
      spacing: { after: 40 },
    }));
    children.push(new D.Paragraph({
      children: [new D.TextRun({
        text: report.icao + " — " + (report.airportName || "") + (report.city ? ", " + report.city : "") + (report.state ? ", " + report.state : ""),
        color: GRAY_TEXT, font: FONT, size: 24,
      })],
      spacing: { after: 160 },
    }));

    children.push(metaTable(D, report));
    children.push(spacer(D));

    children.push(sectionHeading(D, "1.", "DETERMINATION"));
    children.push(determinationTable(D, report.determinations));
    report.determinations.forEach(function (a) {
      a.reasons.forEach(function (r) { children.push(italicNote(D, a.aircraftName + ": " + r)); });
    });
    children.push(spacer(D));

    children.push(sectionHeading(D, "2.", "FACILITY"));
    children.push(italicNote(D, "From FAA NASR. Authoritative."));
    children.push(facilityTable(D, report.facility));
    if (report.facility.remarks.length) {
      children.push(labelPara(D, "NASR REMARKS"));
      children.push(bulletBox(D, report.facility.remarks));
    }
    children.push(spacer(D));

    children.push(sectionHeading(D, "3.", "COMMUNICATIONS"));
    children.push(communicationsTable(D, report.communications));
    children.push(spacer(D));

    children.push(sectionHeading(D, "4.", "RUNWAY EVALUATION"));
    children.push(italicNote(D, "No static minimum runway length applies (PACCAR FOM) -- length/width/weight performance is a function of the day's " +
      "weight, weather, and conditions and belongs in a ForeFlight performance calculation, not a static screen. Declared distances below are a " +
      "reference fact for that calculation, not a pass/fail gate. The only static screen kept here is hard surface + weight-bearing class."));
    groupEndsByRunway(report.endFacts, report.runwayFacts).forEach(function (group) {
      children.push(subHeading(D, "RWY " + group.runwayId + " — " + (group.width ? group.width + " ft wide" : "width NOT AVAILABLE") + ", " + group.surface + ", " + group.wtBrg));
      children.push(declaredDistanceTable(D, group.ends));
      children.push(spacer(D, 60));
    });
    report.determinations.forEach(function (a) {
      children.push(subHeading(D, a.aircraftName + " — hard surface / weight-bearing screen"));
      children.push(runwayScreenTable(D, a.runwayScreens));
      children.push(spacer(D, 60));
    });
    children.push(italicNote(D, "NOT EVALUATED = threshold not yet set in criteria.json. UNKNOWN = threshold set, but NASR has no actual value to compare."));
    children.push(spacer(D));

    children.push(sectionHeading(D, "5.", "NIGHT OPERATIONS & RECENCY"));
    children.push(nightOpsSection(D, report.nightOps));
    children.push(spacer(D));

    children.push(sectionHeading(D, "6.", "SERVICES AND FBO"));
    children.push(banner(D, "NOT FROM NASR — CONFIRM BY PHONE",
      "FBO services, hours, and fuel change without notice. These entries are pilot-entered and/or previously researched -- a starting point and a phone number, not a verified fact.",
      GOLD, BANNER_BG));
    children.push(new D.Paragraph({
      spacing: { after: 100 },
      children: [
        run(D, "Full FBO / services listing for " + report.icao + ": ", { size: 18 }),
        new D.ExternalHyperlink({
          link: report.fboLookupUrl,
          children: [run(D, report.fboLookupUrl, { size: 18, color: NAVY, underline: { type: D.UnderlineType.SINGLE } })],
        }),
      ],
    }));
    children.push(fboTable(D, report.fbo));
    children.push(spacer(D));

    children.push(sectionHeading(D, "7.", "DISQUALIFYING FINDINGS"));
    children.push(report.disqualifyingFindings.length
      ? findingsTable(D, report.disqualifyingFindings)
      : italicNote(D, "None identified from evaluated criteria."));
    children.push(spacer(D));

    children.push(sectionHeading(D, "8.", "ITEMS REQUIRING VERIFICATION"));
    children.push(italicNote(D, "Auto-generated. Every data point the source could not supply, and every marginal screening condition, becomes an action item."));
    children.push(report.verificationItems.length
      ? verificationTable(D, report.verificationItems)
      : italicNote(D, "None — every field used in this report had a NASR-sourced or confirmed value."));
    children.push(spacer(D));

    children.push(sectionHeading(D, "9.", "PIC NARRATIVE"));
    children.push(narrativeBox(D, report.picComments));
    children.push(spacer(D));

    children.push(sectionHeading(D, "10.", "SIGN-OFF"));
    children.push(signOffTable(D, report.picName, report.chiefPilotName));
    children.push(new D.Paragraph({
      children: [new D.TextRun({
        text: "This report is a first-pass screening aid. The Pilot in Command determines actual airport and runway usability for the conditions " +
          "at the time of operation, following a ForeFlight performance calculation. Marginal or flagged items are reviewed with the Chief Pilot per FOM §2.7.4.",
        italics: true, color: GRAY_NOTE, font: FONT, size: 16,
      })],
      spacing: { before: 160 },
      border: { top: { style: D.BorderStyle.SINGLE, size: 4, color: "BFBFBF" } },
    }));

    return new D.Document({
      creator: "Airport Screening Report Generator",
      title: "Airport Screening Report - " + report.icao,
      sections: [{
        properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
        children: children,
      }],
    });
  }

  // -----------------------------------------------------------------------
  // Low-level helpers
  // -----------------------------------------------------------------------
  function run(D, text, opts) {
    opts = opts || {};
    return new D.TextRun(Object.assign({ text: String(text), font: FONT }, opts));
  }
  function spacer(D, size) {
    return new D.Paragraph({ children: [run(D, "", { size: size || 20 })] });
  }
  function sectionHeading(D, num, title) {
    return new D.Paragraph({
      border: { bottom: { style: D.BorderStyle.SINGLE, size: 8, color: NAVY } },
      spacing: { before: 260, after: 120 },
      children: [run(D, num + "  " + title, { bold: true, color: NAVY, size: 24 })],
    });
  }
  function subHeading(D, title) {
    return new D.Paragraph({
      spacing: { before: 160, after: 80 },
      children: [run(D, title, { bold: true, color: NAVY, size: 20 })],
    });
  }
  function labelPara(D, text) {
    return new D.Paragraph({
      spacing: { before: 60, after: 40 },
      children: [run(D, text, { bold: true, color: GRAY_NOTE, size: 16 })],
    });
  }
  function italicNote(D, text) {
    return new D.Paragraph({
      spacing: { after: 100 },
      children: [run(D, text, { italics: true, color: GRAY_NOTE, size: 18 })],
    });
  }
  function cellMargins() {
    return { top: 60, bottom: 60, left: 100, right: 100 };
  }
  function headerCell(D, text, widthPct) {
    return new D.TableCell({
      width: { size: widthPct, type: D.WidthType.PERCENTAGE },
      shading: { fill: NAVY, type: D.ShadingType.CLEAR },
      margins: cellMargins(),
      children: [new D.Paragraph({ children: [run(D, text, { bold: true, color: WHITE, size: 14 })] })],
    });
  }
  function bodyCell(D, text, opts) {
    opts = opts || {};
    return new D.TableCell({
      width: opts.widthPct ? { size: opts.widthPct, type: D.WidthType.PERCENTAGE } : undefined,
      margins: cellMargins(),
      shading: opts.bg ? { fill: opts.bg, type: D.ShadingType.CLEAR } : undefined,
      children: [new D.Paragraph({
        children: [run(D, text === undefined || text === null || text === "" ? "—" : text, {
          bold: !!opts.bold, italics: !!opts.italics, color: opts.color, size: opts.size || 18,
        })],
      })],
    });
  }
  function resultCell(D, result) {
    return bodyCell(D, result, { bold: true, color: RESULT_COLOR[result] || GRAY_TEXT });
  }
  function table(D, rows, columnWidths) {
    return new D.Table({
      width: { size: CONTENT_WIDTH, type: D.WidthType.DXA },
      columnWidths: columnWidths,
      borders: {
        top: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
        bottom: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
        left: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
        right: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
        insideHorizontal: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
        insideVertical: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
      },
      rows: rows,
    });
  }
  function banner(D, title, sub, titleColor, bg) {
    return new D.Table({
      width: { size: CONTENT_WIDTH, type: D.WidthType.DXA },
      borders: {
        top: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
        bottom: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
        left: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
        right: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
        insideHorizontal: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
        insideVertical: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
      },
      rows: [new D.TableRow({
        children: [new D.TableCell({
          shading: { fill: bg, type: D.ShadingType.CLEAR },
          margins: { top: 100, bottom: 100, left: 120, right: 120 },
          children: [
            new D.Paragraph({ children: [run(D, title, { bold: true, color: titleColor, size: 20 })] }),
            new D.Paragraph({ children: [run(D, sub, { color: titleColor === RED ? GOLD : GRAY_TEXT, size: 16 })] }),
          ],
        })],
      })],
    });
  }

  // -----------------------------------------------------------------------
  // Section builders
  // -----------------------------------------------------------------------
  function metaTable(D, report) {
    var headers = ["DATA SOURCE", "ASSESSED", "ICAO", "AIRPORT"];
    var values = ["FAA NASR (28-Day Subscription)", new Date().toISOString().slice(0, 10), report.icao, report.airportName || "—"];
    return table(D, [
      new D.TableRow({ children: headers.map(function (h) { return headerCell(D, h, 25); }) }),
      new D.TableRow({ children: values.map(function (v) { return bodyCell(D, v, { widthPct: 25 }); }) }),
    ]);
  }

  function determinationTable(D, determinations) {
    var headerRow = new D.TableRow({
      children: [headerCell(D, "AIRCRAFT", 30), headerCell(D, "SCREENING RESULT", 70)],
    });
    var rows = determinations.map(function (a) {
      return new D.TableRow({
        children: [
          bodyCell(D, a.aircraftName, { bold: true }),
          bodyCell(D, a.bucket, { bold: true, color: BUCKET_COLOR[a.bucket] }),
        ],
      });
    });
    return table(D, [headerRow].concat(rows));
  }

  function facilityTable(D, facility) {
    function row(label1, val1, label2, val2) {
      return new D.TableRow({
        children: [
          bodyCell(D, label1, { bold: true, size: 14, color: GRAY_NOTE, widthPct: 20 }),
          bodyCell(D, val1, { widthPct: 30 }),
          bodyCell(D, label2, { bold: true, size: 14, color: GRAY_NOTE, widthPct: 20 }),
          bodyCell(D, val2, { widthPct: 30 }),
        ],
      });
    }
    return table(D, [
      row("CONTROL TOWER", facility.controlTower.text, "ELEVATION", facility.elevation),
      row("AIRPORT ATTENDANCE", facility.attendance, "ARFF INDEX", facility.arffIndex),
      row("FUEL ON FIELD", facility.fuelOnField, "AIRPORT MANAGER", facility.airportManagerPhone),
    ]);
  }

  function bulletBox(D, lines) {
    return new D.Table({
      width: { size: CONTENT_WIDTH, type: D.WidthType.DXA },
      borders: {
        top: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" }, bottom: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
        left: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" }, right: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
      },
      rows: [new D.TableRow({
        children: [new D.TableCell({
          margins: { top: 100, bottom: 100, left: 120, right: 120 },
          children: lines.map(function (l) {
            return new D.Paragraph({ bullet: { level: 0 }, children: [run(D, l, { size: 18 })], spacing: { after: 40 } });
          }),
        })],
      })],
    });
  }

  function communicationsTable(D, comms) {
    var headerRow = new D.TableRow({
      children: [headerCell(D, "FACILITY", 22), headerCell(D, "FREQUENCY", 15), headerCell(D, "USE", 20), headerCell(D, "NOTES", 43)],
    });
    if (!comms.length) {
      return table(D, [headerRow, new D.TableRow({ children: [bodyCell(D, "No frequencies on file", { italics: true }), bodyCell(D, "—"), bodyCell(D, "—"), bodyCell(D, "—")] })]);
    }
    var rows = comms.map(function (c) {
      return new D.TableRow({
        children: [
          bodyCell(D, c.call ? c.facilityType + " — " + c.call : c.facilityType),
          bodyCell(D, c.freq),
          bodyCell(D, c.use),
          bodyCell(D, c.remark || "—"),
        ],
      });
    });
    return table(D, [headerRow].concat(rows));
  }

  // Groups the flat per-end fact list back into per-runway blocks, joined
  // with the runway-level physical facts (width/surface/wt brg).
  function groupEndsByRunway(endFacts, runwayFacts) {
    var order = [];
    var byId = {};
    endFacts.forEach(function (ef) {
      if (!byId[ef.runwayId]) { byId[ef.runwayId] = { runwayId: ef.runwayId, ends: [] }; order.push(ef.runwayId); }
      byId[ef.runwayId].ends.push(ef);
    });
    order.forEach(function (id) {
      var rwFact = runwayFacts.filter(function (r) { return r.id === id; })[0];
      if (rwFact) { byId[id].width = rwFact.width; byId[id].surface = rwFact.surface; byId[id].wtBrg = rwFact.wtBrg; }
    });
    return order.map(function (id) { return byId[id]; });
  }

  function declaredDistanceTable(D, ends) {
    var headerRow = new D.TableRow({
      children: [headerCell(D, "END", 14), headerCell(D, "TORA", 20), headerCell(D, "TODA", 20), headerCell(D, "ASDA", 20), headerCell(D, "LDA", 20), headerCell(D, "DISPL THR", 6)],
    });
    var rows = ends.map(function (ef) {
      function distCell(v) { return bodyCell(D, v ? Number(v).toLocaleString("en-US") + " ft" : "NOT AVAILABLE", { color: v ? undefined : GOLD }); }
      return new D.TableRow({
        children: [
          bodyCell(D, ef.endId, { bold: true }),
          distCell(ef.tora), distCell(ef.toda), distCell(ef.asda), distCell(ef.lda),
          bodyCell(D, ef.hasDisplacedThr ? ef.displacedThrLen + " ft" : "—"),
        ],
      });
    });
    return table(D, [headerRow].concat(rows));
  }

  function runwayScreenTable(D, runwayScreens) {
    var headerRow = new D.TableRow({
      children: [headerCell(D, "RUNWAY", 16), headerCell(D, "HARD SURFACE", 28), headerCell(D, "WEIGHT BEARING (DW)", 28), headerCell(D, "VERDICT", 28)],
    });
    var rows = runwayScreens.map(function (rs) {
      var s = rs.rows.filter(function (r) { return r.id === "surface"; })[0];
      var w = rs.rows.filter(function (r) { return r.id === "wtBearingDW"; })[0];
      return new D.TableRow({
        children: [
          bodyCell(D, rs.runwayId, { bold: true }),
          bodyCell(D, s.actualDisplay + " (" + s.result + ")", { color: RESULT_COLOR[s.result] }),
          bodyCell(D, w.actualDisplay + " (" + w.result + ")", { color: RESULT_COLOR[w.result] }),
          resultCell(D, rs.verdict),
        ],
      });
    });
    return table(D, [headerRow].concat(rows));
  }

  function nightOpsSection(D, nightOps) {
    var rows = [];
    function row(label, value, opts) {
      rows.push(new D.TableRow({
        children: [
          bodyCell(D, label, { bold: true, size: 14, color: GRAY_NOTE, widthPct: 35 }),
          bodyCell(D, value, Object.assign({ widthPct: 65 }, opts || {})),
        ],
      }));
    }
    row("2-YEAR RECENCY", nightOps.recency.note);
    row("INSTRUMENT APPROACH", nightOps.instrumentApproach.effectiveLabel +
      (nightOps.instrumentApproach.manualEntryRequired ? " (manual entry" + (nightOps.instrumentApproach.overridden ? ", overridden" : "") + ")" : nightOps.instrumentApproach.overridden ? " (override of auto-derived " + nightOps.instrumentApproach.autoLabel + ")" : " (auto-derived from NASR)"),
      { color: nightOps.instrumentApproach.effective === "unknown" ? GOLD : undefined, bold: nightOps.instrumentApproach.manualEntryRequired });
    row("ALTERNATE MINIMUMS", nightOps.alternateMinimums
      ? nightOps.alternateMinimums.ceilingFt + "-" + nightOps.alternateMinimums.visibilityMi + " (" + nightOps.alternateMinimums.citation + ")"
      : (nightOps.vfrOnlyFlag ? "VFR only -- no instrument alternate minimums apply" : "Cannot be determined until instrument approach availability is confirmed"),
      { color: nightOps.alternateMinimums ? undefined : GOLD });
    row("RUNWAY EDGE LIGHTING (FOM §6.1.6.7)", nightOps.nightCapability.edgeLightingPresent ? "Present" : "Not present / not in NASR",
      { color: nightOps.nightCapability.edgeLightingPresent ? undefined : AMBER });
    row("LIGHTED WIND INDICATOR (FOM §6.1.6.7)", nightOps.nightCapability.windIndicatorLighted ? "Present" : "Not present / not in NASR",
      { color: nightOps.nightCapability.windIndicatorLighted ? undefined : AMBER });
    row("OBSTRUCTION LIGHTING (FOM §6.1.6.7)", nightOps.nightCapability.obstructionLightingPresent ? "Present (marked/lighted obstruction on file)" : "None on file",
      { color: nightOps.nightCapability.obstructionLightingPresent ? undefined : AMBER });
    row("NIGHT CIRCLING", nightOps.nightCirclingNote, { color: AMBER });
    row("REMOTE / UNMONITORED (FOM §6.1.6.7)", nightOps.remoteUnmonitored
      ? "YES -- uncontrolled, no on-field weather reporting" + (nightOps.winterNightHardFlag ? " -- HARD FLAG: planned winter-season night operation" : "")
      : "No -- towered or on-field weather reporting available",
      { color: nightOps.winterNightHardFlag ? RED : nightOps.remoteUnmonitored ? AMBER : undefined, bold: nightOps.winterNightHardFlag });
    row("WINTER NIGHT OPERATION PLANNED", nightOps.winterNightOps === "yes" ? "Yes" : nightOps.winterNightOps === "no" ? "No" : "Unknown / not specified",
      { color: nightOps.winterNightOps === "unknown" ? GOLD : undefined });
    return table(D, rows);
  }

  function fboTable(D, fbo) {
    var headerRow = new D.TableRow({
      children: [headerCell(D, "FBO", 25), headerCell(D, "PHONE", 20), headerCell(D, "FUEL", 25), headerCell(D, "SERVICES", 30)],
    });
    if (!fbo.length) {
      return table(D, [headerRow, new D.TableRow({
        children: [bodyCell(D, "No FBO records on file for this airport", { italics: true, color: GOLD }), bodyCell(D, "—"), bodyCell(D, "—"), bodyCell(D, "—")],
      })]);
    }
    var rows = fbo.map(function (f) {
      return new D.TableRow({
        children: [
          bodyCell(D, f.name, { bold: true }),
          bodyCell(D, f.phone, { color: f.phone === "CONFIRM" ? GOLD : undefined, bold: f.phone === "CONFIRM" }),
          bodyCell(D, f.fuel, { color: f.fuel === "CONFIRM" ? GOLD : undefined, bold: f.fuel === "CONFIRM" }),
          bodyCell(D, f.services, { color: f.services === "CONFIRM" ? GOLD : undefined, bold: f.services === "CONFIRM" }),
        ],
      });
    });
    return table(D, [headerRow].concat(rows));
  }

  function findingsTable(D, findings) {
    return table(D, findings.map(function (f) {
      return new D.TableRow({ children: [bodyCell(D, f, { color: RED })] });
    }));
  }

  function verificationTable(D, items) {
    var headerRow = new D.TableRow({ children: [headerCell(D, "#", 6), headerCell(D, "ITEM", 52), headerCell(D, "ACTION", 42)] });
    var rows = items.map(function (it, i) {
      return new D.TableRow({ children: [bodyCell(D, String(i + 1)), bodyCell(D, it.item), bodyCell(D, it.action || "Confirm before dispatch")] });
    });
    return table(D, [headerRow].concat(rows));
  }

  // PIC-typed free text if provided; otherwise blank lines for a hand-
  // written narrative when the report is printed.
  function narrativeBox(D, text) {
    var paras;
    if (text && text.trim()) {
      paras = text.split("\n").map(function (line) {
        return new D.Paragraph({ children: [run(D, line, { size: 20 })], spacing: { after: 60 } });
      });
    } else {
      paras = [0, 1, 2].map(function () { return new D.Paragraph({ children: [run(D, "", { size: 20 })] }); });
    }
    return new D.Table({
      width: { size: CONTENT_WIDTH, type: D.WidthType.DXA },
      borders: {
        top: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" }, bottom: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
        left: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" }, right: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
      },
      rows: [new D.TableRow({ children: [new D.TableCell({ margins: { top: 200, bottom: 200, left: 120, right: 120 }, children: paras })] })],
    });
  }

  function signOffTable(D, picName, chiefPilotName) {
    var headerRow = new D.TableRow({ children: [headerCell(D, "ROLE", 25), headerCell(D, "NAME", 30), headerCell(D, "SIGNATURE", 25), headerCell(D, "DATE", 20)] });
    var rows = [["Pilot in Command", picName || ""], ["Chief Pilot", chiefPilotName || ""]].map(function (r) {
      return new D.TableRow({ children: [bodyCell(D, r[0], { bold: true }), bodyCell(D, r[1]), bodyCell(D, ""), bodyCell(D, "")] });
    });
    return table(D, [headerRow].concat(rows));
  }

  return { build: build };
});
