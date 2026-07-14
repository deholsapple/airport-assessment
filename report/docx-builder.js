// Renders a report-data.js report model into a docx.Document, matching the
// layout of KDTO-suitability-SAMPLE.docx (colors, table structure, section
// order). Takes the docx library as a parameter (window.docx in the browser,
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
  var RED = "C00000";
  var GOLD = "7F6000";
  var AMBER = "BF8F00";
  var GREEN = "375623";
  var GRAY_TEXT = "404040";
  var GRAY_NOTE = "595959";
  var WHITE = "FFFFFF";

  var CONTENT_WIDTH = 9360; // 6.5in at 1440 dxa/in -- matches the sample's page margins.
  var FONT = "Calibri";

  var RESULT_COLOR = {
    PASS: GREEN, CAUTION: AMBER, FAIL: RED, UNKNOWN: GOLD, "NOT EVALUATED": GRAY_NOTE,
  };
  var DETERMINATION_COLOR = {
    SUITABLE: GREEN, "SUITABLE WITH LIMITATIONS": AMBER, "NOT SUITABLE": RED,
  };

  function build(report, D) {
    var children = [];

    if (!report.verified) {
      children.push(banner(D,
        "CRITERIA NOT VERIFIED",
        "One or more evaluation thresholds in criteria.json have not been reviewed against the AFM / ops specs. " +
        "Every criterion below that depends on an unset threshold reads NOT EVALUATED. Do not use for dispatch.",
        RED));
    }

    children.push(new D.Paragraph({
      children: [new D.TextRun({ text: "AIRPORT SUITABILITY ASSESSMENT", bold: true, color: NAVY, font: FONT, size: 32 })],
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
      if (a.note) children.push(italicNote(D, a.note));
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

    children.push(sectionHeading(D, "4.", "RUNWAY EVALUATION — ALL RUNWAYS"));
    children.push(runwayEvalTable(D, report));
    children.push(spacer(D));

    children.push(sectionHeading(D, "5.", "CRITERIA EVALUATION"));
    var anyCriteriaTable = false;
    report.determinations.forEach(function (a) {
      a.criteriaTables.forEach(function (t) {
        anyCriteriaTable = true;
        children.push(subHeading(D, a.aircraftName + " / RWY " + t.runwayId));
        children.push(criteriaTable(D, t.rows));
        children.push(spacer(D, 60));
      });
    });
    if (!anyCriteriaTable) {
      children.push(italicNote(D, "No runway currently qualifies for a detailed criteria breakdown for any evaluated aircraft."));
    }
    children.push(italicNote(D, "NOT EVALUATED = threshold not yet set in criteria.json. UNKNOWN = threshold set, but NASR has no actual value to compare."));
    children.push(spacer(D));

    children.push(sectionHeading(D, "6.", "SERVICES AND FBO"));
    children.push(banner(D, "NOT FROM NASR — CONFIRM BY PHONE",
      "FBO services, hours, and fuel change without notice. These entries are a starting point and a phone number, not a verified fact.",
      GOLD));
    children.push(fboTable(D, report.fbo));
    children.push(spacer(D));

    children.push(sectionHeading(D, "7.", "DISQUALIFYING FINDINGS"));
    children.push(report.disqualifyingFindings.length
      ? findingsTable(D, report.disqualifyingFindings)
      : italicNote(D, "None identified from evaluated criteria."));
    children.push(spacer(D));

    children.push(sectionHeading(D, "8.", "ITEMS REQUIRING VERIFICATION"));
    children.push(italicNote(D, "Auto-generated. Every data point the source could not supply becomes an action item."));
    children.push(report.verificationItems.length
      ? verificationTable(D, report.verificationItems)
      : italicNote(D, "None — every field used in this report had a NASR-sourced or confirmed value."));
    children.push(spacer(D));

    children.push(sectionHeading(D, "9.", "PIC NARRATIVE"));
    children.push(blankBox(D, 3));
    children.push(spacer(D));

    children.push(sectionHeading(D, "10.", "SIGN-OFF"));
    children.push(signOffTable(D));
    children.push(new D.Paragraph({
      children: [new D.TextRun({
        text: "This assessment is a planning guide. The Pilot in Command determines actual airport and runway usability for the conditions at the time of operation.",
        italics: true, color: GRAY_NOTE, font: FONT, size: 16,
      })],
      spacing: { before: 160 },
      border: { top: { style: D.BorderStyle.SINGLE, size: 4, color: "BFBFBF" } },
    }));

    return new D.Document({
      creator: "Airport Suitability Report Generator",
      title: "Airport Suitability Assessment - " + report.icao,
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
  function banner(D, title, sub, color) {
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
          shading: { fill: BANNER_BG, type: D.ShadingType.CLEAR },
          margins: { top: 100, bottom: 100, left: 120, right: 120 },
          children: [
            new D.Paragraph({ children: [run(D, title, { bold: true, color: color, size: 20 })] }),
            new D.Paragraph({ children: [run(D, sub, { color: GOLD, size: 16 })] }),
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
      children: [headerCell(D, "AIRCRAFT", 30), headerCell(D, "DETERMINATION", 35), headerCell(D, "QUALIFYING RUNWAYS", 35)],
    });
    var rows = determinations.map(function (a) {
      var qual = a.qualifyingRunways.length
        ? a.qualifyingRunways.join(", ") + (a.qualifyingRunways.length === 1 ? " only" : "")
        : "None";
      return new D.TableRow({
        children: [
          bodyCell(D, a.aircraftName, { bold: true }),
          bodyCell(D, a.determination, { bold: true, color: DETERMINATION_COLOR[a.determination] }),
          bodyCell(D, qual),
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

  function runwayEvalTable(D, report) {
    var aircraftNames = report.determinations.map(function (a) { return a.aircraftName; });
    var headers = ["RUNWAY", "LENGTH", "WIDTH", "SURFACE", "WT BRG"].concat(aircraftNames);
    var headerRow = new D.TableRow({ children: headers.map(function (h) { return headerCell(D, h, 100 / headers.length); }) });
    var rows = report.runwayFacts.map(function (rw, i) {
      var cells = [
        bodyCell(D, rw.id, { bold: true }),
        bodyCell(D, rw.length ? rw.length.toLocaleString("en-US") + " ft" : "—"),
        bodyCell(D, rw.width ? rw.width + " ft" : "—"),
        bodyCell(D, rw.surface),
        bodyCell(D, rw.wtBrg),
      ];
      report.determinations.forEach(function (a) {
        var h = a.headline[i];
        var text = h.result + (h.marginDisplay && h.marginDisplay !== "—" ? " (" + h.marginDisplay + ")" : "");
        cells.push(bodyCell(D, text, { bold: true, color: RESULT_COLOR[h.result] }));
      });
      return new D.TableRow({ children: cells });
    });
    return table(D, [headerRow].concat(rows));
  }

  function criteriaTable(D, rows) {
    var headerRow = new D.TableRow({
      children: [headerCell(D, "CRITERION", 26), headerCell(D, "ACTUAL", 22), headerCell(D, "REQUIRED", 20), headerCell(D, "MARGIN", 14), headerCell(D, "RESULT", 18)],
    });
    var dataRows = rows.map(function (r) {
      return new D.TableRow({
        children: [
          bodyCell(D, r.label, { bold: true }),
          bodyCell(D, r.actualDisplay, { color: r.actualDisplay === "NOT AVAILABLE" ? GOLD : undefined }),
          bodyCell(D, r.requiredDisplay),
          bodyCell(D, r.marginDisplay),
          resultCell(D, r.result),
        ],
      });
    });
    return table(D, [headerRow].concat(dataRows));
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

  function blankBox(D, lines) {
    var paras = [];
    for (var i = 0; i < lines; i++) paras.push(new D.Paragraph({ children: [run(D, "", { size: 20 })] }));
    return new D.Table({
      width: { size: CONTENT_WIDTH, type: D.WidthType.DXA },
      borders: {
        top: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" }, bottom: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
        left: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" }, right: { style: D.BorderStyle.SINGLE, size: 4, color: "auto" },
      },
      rows: [new D.TableRow({ children: [new D.TableCell({ margins: { top: 200, bottom: 200, left: 120, right: 120 }, children: paras })] })],
    });
  }

  function signOffTable(D) {
    var headerRow = new D.TableRow({ children: [headerCell(D, "ROLE", 25), headerCell(D, "NAME", 30), headerCell(D, "SIGNATURE", 25), headerCell(D, "DATE", 20)] });
    var rows = ["Pilot in Command", "Chief Pilot"].map(function (role) {
      return new D.TableRow({ children: [bodyCell(D, role, { bold: true }), bodyCell(D, ""), bodyCell(D, ""), bodyCell(D, "")] });
    });
    return table(D, [headerRow].concat(rows));
  }

  return { build: build };
});
