// Pure computation layer for the airport suitability report.
//
// Takes NASR airport data + criteria.json + the FBO directory and produces
// a plain-JS "report model" -- no DOM, no docx library -- so it can run
// identically in the browser (index.html) and under Node (verify-report.js).
//
// Core honesty rules this module enforces everywhere, not just in one spot:
//   - A criterion with threshold === null always reads NOT EVALUATED, no
//     matter what NASR data is available for it.
//   - A criterion with a real threshold but no actual data reads UNKNOWN.
//   - Any value this module cannot source from NASR renders as the literal
//     string "NOT AVAILABLE" (never a blank, never a guess), and is also
//     logged as an "Item Requiring Verification".
//   - Any FBO fact not present in fbo-data.js renders as "CONFIRM" and is
//     also logged as a verification item.
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.ReportData = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var NOT_AVAILABLE = "NOT AVAILABLE";
  var CONFIRM = "CONFIRM";

  var ARFF_INDEX_ORDER = ["A", "B", "C", "D", "E"];

  var RWY_LIGHT_ACRONYM = {
    HIGH: "HIRL", MED: "MIRL", LOW: "LIRL", NSTD: "NSTD", PERI: "PERI", STRB: "STRB", FLD: "FLD",
  };

  var FUEL_SHORT_LABEL = {
    "80": "Grade 80", "100": "Grade 100", "100LL": "100LL", "115": "Grade 115",
    A: "Jet A", "A+": "Jet A+", "A++": "Jet A++", "A++10": "Jet A++10",
    A1: "Jet A-1", "A1+": "Jet A-1+", B: "Jet B", "B+": "Jet B+",
    J: "Jet (unspec.)", J4: "JP-4", J5: "JP-5", J8: "JP-8", "J8+10": "JP-8+10",
    MOGAS: "Mogas", UL91: "UL91", UL94: "UL94",
  };

  // -----------------------------------------------------------------------
  // Small helpers
  // -----------------------------------------------------------------------
  function fmtNum(n) {
    return n.toLocaleString("en-US");
  }
  function fmtSigned(n, unit) {
    var sign = n >= 0 ? "+" : "-";
    return sign + fmtNum(Math.abs(n)) + (unit ? " " + unit : "");
  }
  function titleCase(s) {
    if (!s) return "";
    return String(s).replace(/\w\S*/g, function (t) {
      return t.charAt(0) + t.slice(1).toLowerCase();
    });
  }
  function parseIntOrNull(v) {
    if (v === undefined || v === null || v === "") return null;
    var n = parseInt(v, 10);
    return isFinite(n) ? n : null;
  }

  // Verification-item accumulator. Every "gap" the tool renders (NOT
  // AVAILABLE from NASR, or CONFIRM on an unverified FBO fact) is logged
  // here exactly once (by item text) so Section 8 can be auto-generated
  // straight from what the rest of the report actually rendered.
  function makeGapTracker() {
    var seen = {};
    var items = [];
    return {
      na: function (item, action) {
        if (!seen[item]) { seen[item] = true; items.push({ item: item, action: action }); }
        return NOT_AVAILABLE;
      },
      confirm: function (item, action) {
        if (!seen[item]) { seen[item] = true; items.push({ item: item, action: action }); }
        return CONFIRM;
      },
      items: items,
    };
  }

  // -----------------------------------------------------------------------
  // Facility (Section 2)
  // -----------------------------------------------------------------------
  function buildFacility(rec, gaps) {
    var tower = rec.tower
      ? { present: true, text: "YES" + (rec.tower.hours ? " — " + rec.tower.hours + " L" : "") }
      : { present: false, text: "NO" };

    var attendance;
    if (!rec.attendance || rec.attendance.length === 0) {
      attendance = gaps.na("Airport attendance schedule not in NASR", "Confirm hours with airport ops" + (managerPhone(rec) ? " — " + managerPhone(rec) : ""));
    } else {
      attendance = rec.attendance
        .map(function (seg) {
          var prefix = seg.month && seg.month !== "ALL" ? seg.month + ": " : "";
          var days = seg.days && seg.days !== "ALL" ? seg.days + " " : "";
          return prefix + days + (seg.hours || "");
        })
        .join("; ");
    }

    var arff = arffIndexOf(rec, gaps);

    var fuel = rec.fuelTypesDesc && rec.fuelTypesDesc.length
      ? rec.fuelTypesDesc.map(function (f) { return FUEL_SHORT_LABEL[f.code] || f.code; }).join(", ")
      : gaps.na("Fuel type not in NASR", "Confirm fuel availability with airport ops");

    var mgrPhone = managerPhone(rec) || gaps.na("Airport manager phone not in NASR", "Obtain from airport ops / Chart Supplement");

    return {
      controlTower: tower,
      elevation: rec.elev !== null && rec.elev !== undefined ? fmtNum(rec.elev) + " ft MSL" : gaps.na("Field elevation not in NASR", "Confirm via Chart Supplement"),
      attendance: attendance,
      arffIndex: arff.display,
      fuelOnField: fuel,
      airportManagerPhone: mgrPhone,
      remarks: rec.remarks && rec.remarks.length ? rec.remarks.slice() : [],
    };
  }

  function managerPhone(rec) {
    var mgr = (rec.contacts || []).find(function (c) { return /manager/i.test(c.title || ""); });
    return mgr && mgr.phone ? mgr.phone : "";
  }

  // FAR_139_TYPE_CODE looks like "I B" (Class I, ARFF Index B). Blank means
  // the airport isn't FAR Part 139 certificated -- there is no ARFF index
  // to report, not a data gap NASR failed to fill.
  function arffIndexOf(rec, gaps) {
    var raw = (rec.far139Type || "").trim();
    if (!raw) {
      return { letter: null, display: gaps.na("ARFF index not published in NASR (not a FAR Part 139 airport)", "Confirm with airport ops" + (managerPhone(rec) ? " — " + managerPhone(rec) : "")) };
    }
    var parts = raw.split(/\s+/);
    var letter = parts[parts.length - 1];
    if (ARFF_INDEX_ORDER.indexOf(letter) === -1) {
      return { letter: null, display: raw };
    }
    return { letter: letter, display: "Index " + letter + " (FAR 139 Class " + parts[0] + ")" };
  }

  // -----------------------------------------------------------------------
  // Communications (Section 3)
  // -----------------------------------------------------------------------
  function buildCommunications(rec) {
    return (rec.frequencies || []).map(function (f) {
      return {
        facilityType: f.facilityType,
        call: f.call,
        freq: f.freq,
        use: f.use,
        remark: f.remark,
      };
    });
  }

  // -----------------------------------------------------------------------
  // Runway lighting summary used by both Section 4 (terse) and criteria (Section 5)
  // -----------------------------------------------------------------------
  function lightingSummary(rw) {
    var parts = [];
    if (rw.lighting) parts.push(RWY_LIGHT_ACRONYM[rw.lighting] || rw.lighting);
    var apchCodes = [];
    (rw.ends || []).forEach(function (e) {
      if (e.apchLgtSystem && apchCodes.indexOf(e.apchLgtSystem) === -1) apchCodes.push(e.apchLgtSystem);
    });
    if (apchCodes.length) parts.push(apchCodes.join("/"));
    return parts.join(" + ");
  }

  function weightBearingCompact(rw) {
    var parts = [];
    if (rw.grossWtSW) parts.push("S-" + rw.grossWtSW);
    if (rw.grossWtDW) parts.push("D-" + rw.grossWtDW);
    if (rw.grossWtDTW) parts.push("DT-" + rw.grossWtDTW);
    if (rw.grossWtDDTW) parts.push("DDT-" + rw.grossWtDDTW);
    return parts.join(" ");
  }

  // -----------------------------------------------------------------------
  // Criterion evaluation -- the single place PASS/CAUTION/FAIL/UNKNOWN/
  // NOT EVALUATED gets decided. Threshold-null always wins: it means the
  // criterion has never been reviewed against the AFM/ops specs, so no
  // comparison is attempted no matter what actual data exists.
  // -----------------------------------------------------------------------
  function evaluateCriterion(def, actual) {
    var row = {
      id: def.id, label: def.label, unit: def.unit,
      actualDisplay: null, requiredDisplay: null, marginDisplay: "—", result: null,
    };

    var thresholdMissing =
      def.compare === "categorical" ? (def.allowed === null || def.allowed === undefined) : (def.threshold === null || def.threshold === undefined);

    if (thresholdMissing) {
      row.result = "NOT EVALUATED";
      row.requiredDisplay = "—";
      row.actualDisplay = actual.display;
      return row;
    }

    if (!actual.available) {
      row.result = "UNKNOWN";
      row.actualDisplay = actual.display;
      row.requiredDisplay = requiredDisplayFor(def);
      return row;
    }

    row.actualDisplay = actual.display;
    row.requiredDisplay = requiredDisplayFor(def);

    if (def.compare === "min") {
      var marginMin = actual.value - def.threshold;
      row.marginDisplay = fmtSigned(marginMin, def.unit);
      row.result = marginMin >= 0 ? "PASS" : "FAIL";
      if (row.result === "PASS" && def.cautionBand && marginMin < def.cautionBand) row.result = "CAUTION";
    } else if (def.compare === "max") {
      var marginMax = def.threshold - actual.value;
      row.marginDisplay = fmtSigned(marginMax, def.unit);
      row.result = marginMax >= 0 ? "PASS" : "FAIL";
      if (row.result === "PASS" && def.cautionBand && marginMax < def.cautionBand) row.result = "CAUTION";
    } else if (def.compare === "boolean") {
      row.result = actual.value === true ? "PASS" : "FAIL";
    } else if (def.compare === "categorical") {
      row.result = def.allowed.indexOf(actual.value) !== -1 ? "PASS" : "FAIL";
    } else if (def.compare === "ordinal") {
      var actualRank = ARFF_INDEX_ORDER.indexOf(actual.value);
      var reqRank = ARFF_INDEX_ORDER.indexOf(def.threshold);
      row.result = actualRank >= reqRank ? "PASS" : "FAIL";
    }
    return row;
  }

  function requiredDisplayFor(def) {
    if (def.compare === "categorical") return def.requiredLabel || def.allowed.join(", ");
    if (def.compare === "boolean") return def.threshold ? "Required" : "—";
    if (def.compare === "ordinal") return "Index " + def.threshold;
    var unit = def.unit ? " " + def.unit : "";
    if (def.compare === "min") return fmtNum(def.threshold) + unit;
    return "≤ " + fmtNum(def.threshold) + unit;
  }

  // Resolves the "actual" side for one criterion against one runway (plus
  // airport-level facts that don't vary by runway).
  function actualFor(criterionId, rec, rw, gaps) {
    switch (criterionId) {
      case "rwyLength":
        return rw.length ? { available: true, value: rw.length, display: fmtNum(rw.length) + " ft" }
          : { available: false, display: gaps.na("Runway " + rw.id + " length not in NASR", "Verify via Chart Supplement") };
      case "rwyWidth":
        return rw.width ? { available: true, value: rw.width, display: fmtNum(rw.width) + " ft" }
          : { available: false, display: gaps.na("Runway " + rw.id + " width not in NASR", "Verify via Chart Supplement") };
      case "wtBearingDW": {
        var dw = parseIntOrNull(rw.grossWtDW);
        if (dw === null) return { available: false, display: gaps.na("Runway " + rw.id + " dual-wheel weight bearing not in NASR", "Verify PCN/gross weight via Chart Supplement") };
        return { available: true, value: dw * 1000, display: fmtNum(dw * 1000) + " lb" };
      }
      case "surface": {
        if (!rw.surface) return { available: false, display: gaps.na("Runway " + rw.id + " surface not in NASR", "Verify via Chart Supplement") };
        var baseToken = rw.surface.split(/[-/]/)[0];
        return { available: true, value: baseToken, display: rw.surfaceDesc || rw.surface };
      }
      case "lighting": {
        var summary = lightingSummary(rw);
        var hasLighting = !!(rw.lighting && rw.lighting !== "NONE");
        return { available: true, value: hasLighting, display: summary || "None" };
      }
      case "arffIndex": {
        var arff = arffIndexOf(rec, gaps);
        return arff.letter
          ? { available: true, value: arff.letter, display: arff.display }
          : { available: false, display: arff.display };
      }
      case "apchMinimums":
        return { available: false, display: gaps.na("Published approach minimums not in NASR (IAP/TERPS data, not an airport master record)", "Verify via current IAP charts") };
      case "alternateDistance":
        return { available: false, display: gaps.na("Alternate airport selection not automated", "PIC/dispatch to select and verify a qualifying alternate") };
      default:
        return { available: false, display: NOT_AVAILABLE };
    }
  }

  // -----------------------------------------------------------------------
  // Runway evaluation (Section 4) + criteria evaluation (Section 5) +
  // determination (Section 1), computed together since determination
  // depends on the full per-runway criteria results.
  // -----------------------------------------------------------------------
  function buildRunwaysAndCriteria(rec, criteriaData, gaps) {
    var runwayFacts = (rec.runways || []).map(function (rw) {
      return {
        id: rw.id,
        length: rw.length,
        width: rw.width,
        surface: rw.surfaceDesc || rw.surface || gaps.na("Runway " + rw.id + " surface not in NASR", "Verify via Chart Supplement"),
        wtBrg: weightBearingCompact(rw) || gaps.na("Runway " + rw.id + " weight bearing not in NASR", "Verify via Chart Supplement"),
      };
    });

    var perAircraft = criteriaData.aircraft.map(function (aircraft) {
      var lengthDef = aircraft.criteria.filter(function (c) { return c.id === "rwyLength"; })[0];

      // Section 4: headline per-runway verdict, driven by runway length only.
      var headline = (rec.runways || []).map(function (rw) {
        var actual = actualFor("rwyLength", rec, rw, gaps);
        var row = evaluateCriterion(lengthDef, actual);
        return { runwayId: rw.id, result: row.result, marginDisplay: row.marginDisplay };
      });

      // Section 5: full multi-criterion table, one per runway that isn't a
      // headline FAIL (a runway that fails on raw length isn't "qualifying"
      // and doesn't get a detailed breakdown -- matches the sample).
      var criteriaTables = [];
      (rec.runways || []).forEach(function (rw, i) {
        if (headline[i].result === "FAIL") return;
        var rows = aircraft.criteria.map(function (def) {
          var actual = actualFor(def.id, rec, rw, gaps);
          return evaluateCriterion(def, actual);
        });
        var worst = worstResult(rows.map(function (r) { return r.result; }));
        criteriaTables.push({ runwayId: rw.id, rows: rows, verdict: worst });
      });

      // Determination: which runways are usable (no FAIL in the full table)
      // vs. clean (every row PASS) vs. disqualified.
      var qualifying = criteriaTables.filter(function (t) { return t.verdict !== "FAIL"; });
      var clean = qualifying.filter(function (t) { return t.verdict === "PASS"; });

      var determination;
      if (clean.length > 0) determination = "SUITABLE";
      else if (qualifying.length > 0) determination = "SUITABLE WITH LIMITATIONS";
      else determination = "NOT SUITABLE";

      var qualifyingIds = qualifying.map(function (t) { return t.runwayId; });
      var note = null;
      if (qualifying.length === 0) {
        note = aircraft.name + ": no runway meets evaluated criteria (or none have been evaluated yet).";
      } else if (qualifying.length === 1) {
        note = aircraft.name + ": single qualifying runway (" + qualifyingIds[0] + "). If it is out of service, no runway on the field currently qualifies.";
      }

      return {
        aircraftId: aircraft.id,
        aircraftName: aircraft.name,
        headline: headline,
        determination: determination,
        qualifyingRunways: qualifyingIds,
        note: note,
        criteriaTables: criteriaTables,
      };
    });

    return { runwayFacts: runwayFacts, perAircraft: perAircraft };
  }

  var RESULT_SEVERITY = { FAIL: 3, UNKNOWN: 2, CAUTION: 2, "NOT EVALUATED": 1, PASS: 0 };
  function worstResult(results) {
    var worst = "PASS";
    results.forEach(function (r) {
      if (RESULT_SEVERITY[r] > RESULT_SEVERITY[worst]) worst = r;
    });
    return worst;
  }

  // -----------------------------------------------------------------------
  // Services & FBO (Section 6)
  // -----------------------------------------------------------------------
  function buildFbo(icao, fboData, gaps) {
    var list = (fboData && fboData[icao]) || [];
    if (list.length === 0) return [];
    return list.map(function (fbo) {
      return {
        name: fbo.name || gaps.confirm(fbo.name ? "" : "Unnamed FBO record for " + icao, "Identify and confirm FBO name"),
        phone: fbo.phone || gaps.confirm((fbo.name || "FBO") + " phone not confirmed", "Look up and confirm before dispatch"),
        fuel: fbo.fuel || gaps.confirm((fbo.name || "FBO") + " fuel availability not confirmed", "Call to confirm fuel"),
        services: [fbo.gpu ? "GPU" : "", fbo.deice ? "De-ice" : "", fbo.hangar ? "Hangar" : "", fbo.hours ? "Hours: " + fbo.hours : ""]
          .filter(Boolean).join(", ") || fbo.sourceNote || gaps.confirm((fbo.name || "FBO") + " services not confirmed", "Call to confirm services"),
      };
    });
  }

  // -----------------------------------------------------------------------
  // Disqualifying findings (Section 7) -- mechanical, derived only from
  // actual FAIL results already computed above. No new judgment calls.
  // -----------------------------------------------------------------------
  function buildDisqualifyingFindings(rec, perAircraft) {
    var findings = [];
    (rec.runways || []).forEach(function (rw) {
      var failing = perAircraft.filter(function (a) {
        var h = a.headline.filter(function (h) { return h.runwayId === rw.id; })[0];
        return h && h.result === "FAIL";
      });
      if (failing.length === 0) return;
      var names = failing.map(function (a) { return a.aircraftName; });
      var margins = failing.map(function (a) {
        var h = a.headline.filter(function (h) { return h.runwayId === rw.id; })[0];
        return h.marginDisplay;
      });
      var allSame = margins.every(function (m) { return m === margins[0]; });
      var text = allSame
        ? "RWY " + rw.id + " (" + fmtNum(rw.length || 0) + " ft): below minimum length for " + names.join(" and ") + " (" + margins[0] + ")."
        : "RWY " + rw.id + " (" + fmtNum(rw.length || 0) + " ft): below minimum length for " + names.map(function (n, i) { return n + " (" + margins[i] + ")"; }).join("; ") + ".";
      findings.push(text);
    });
    perAircraft.forEach(function (a) {
      if (a.determination === "NOT SUITABLE") {
        findings.push(a.aircraftName + " — airport: no runway meets evaluated criteria for this aircraft.");
      }
    });
    return findings;
  }

  // -----------------------------------------------------------------------
  // Top-level build
  // -----------------------------------------------------------------------
  function buildReport(icao, airportData, criteriaData, fboData) {
    var rec = (airportData || {})[icao];
    if (!rec) return { error: "No NASR record found for ICAO \"" + icao + "\" in the embedded dataset." };

    var gaps = makeGapTracker();

    var facility = buildFacility(rec, gaps);
    var communications = buildCommunications(rec);
    var runwaysAndCriteria = buildRunwaysAndCriteria(rec, criteriaData, gaps);
    var fbo = buildFbo(icao, fboData, gaps);
    var disqualifying = buildDisqualifyingFindings(rec, runwaysAndCriteria.perAircraft);

    return {
      icao: rec.icao,
      airportName: rec.name,
      city: rec.city,
      state: rec.state,
      verified: !!criteriaData.verified,
      criteriaNote: (criteriaData.meta && criteriaData.meta.note) || "",
      facility: facility,
      communications: communications,
      runwayFacts: runwaysAndCriteria.runwayFacts,
      determinations: runwaysAndCriteria.perAircraft,
      fbo: fbo,
      disqualifyingFindings: disqualifying,
      verificationItems: gaps.items,
    };
  }

  return { buildReport: buildReport, NOT_AVAILABLE: NOT_AVAILABLE, CONFIRM: CONFIRM };
});
