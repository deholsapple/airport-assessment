// Pure computation layer for the airport suitability report.
//
// Takes NASR airport data + criteria.json + the FBO directory (plus
// pilot-entered extras: manual FBO rows, PIC comments, sign-off names) and
// produces a plain-JS "report model" -- no DOM, no docx library -- so it
// can run identically in the browser (index.html) and under Node
// (verify-report.js).
//
// Core honesty rules this module enforces everywhere, not just in one spot:
//   - A criterion with threshold === null always reads NOT EVALUATED, no
//     matter what NASR data is available for it.
//   - A criterion with a real threshold but no actual data reads UNKNOWN.
//   - Any value this module cannot source from NASR renders as the literal
//     string "NOT AVAILABLE" (never a blank, never a guess), and is also
//     logged as an "Item Requiring Verification".
//   - Any FBO fact not present in fbo-data.js / not typed in by the pilot
//     renders as "CONFIRM" and is also logged as a verification item.
//   - Runway suitability is evaluated per END against NASR's declared
//     distances (TORA/TODA/ASDA/LDA), never against physical runway
//     length -- a displaced threshold or obstacle can make the declared
//     distance shorter than the physical length, and using length would
//     over-credit the runway in the dangerous direction. A missing
//     declared distance is NOT AVAILABLE, never a fallback to length.
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

  // Communications whitelist (Fix 1). NASR's FRQ.csv lists an airport's
  // TRACON-served named arrival/departure procedure frequencies (one row
  // per STAR/DP/RNAV transition) right alongside its own tower/ground/
  // CTAF/ATIS rows. Those procedure rows are real data but they are FMS
  // routing noise for a suitability report -- a busy Class-B-served field
  // can have 30+ of them burying the one tower frequency a crew needs.
  var NAMED_PROCEDURE_RE = /\b(STAR|DP|RNAV)\b/;
  function isOperationalFrequency(f) {
    var use = (f.use || "").toUpperCase().trim();
    if (!use) return false;
    if (NAMED_PROCEDURE_RE.test(use)) return false; // named arrival/departure procedure -- drop regardless of category
    // UHF military band (225+ MHz) duplicate of a civil VHF frequency --
    // not usable by a civilian business jet, so not "operational" here.
    var freqNum = parseFloat(f.freq);
    if (isFinite(freqNum) && freqNum >= 225) return false;
    if (use === "ATIS" || use === "CTAF") return true;
    if (/\bASOS\b|\bAWOS\b/.test(use)) return true;
    if (/^(LCL|GND|CD)\/P$/.test(use)) return true; // tower / ground / clearance delivery, primary
    if (/^(APCH\/P( DEP\/P)?|DEP\/P)$/.test(use)) return true; // primary approach/departure only (not secondary, not named)
    return false;
  }

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
  // Communications (Section 3) -- Fix 1: operational set only, de-duped.
  // -----------------------------------------------------------------------
  function buildCommunications(rec) {
    var seen = {};
    var result = [];
    (rec.frequencies || []).forEach(function (f) {
      if (!isOperationalFrequency(f)) return;
      var key = (f.freq || "") + "|" + (f.use || "");
      if (seen[key]) return;
      seen[key] = true;
      result.push({ facilityType: f.facilityType, call: f.call, freq: f.freq, use: f.use, remark: f.remark });
    });
    return result;
  }

  // -----------------------------------------------------------------------
  // Runway lighting summary + weight-bearing compact notation, shared by
  // the merged runway/criteria section.
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

  // Resolves the "actual" side for one criterion against one runway end
  // (ctx.end) plus its parent runway (ctx.rw) and the airport (rec).
  function actualFor(criterionId, rec, ctx, gaps) {
    var rw = ctx.rw, end = ctx.end;
    var label = end.id + " (RWY " + rw.id + ")";
    switch (criterionId) {
      case "takeoffDistance": {
        var tora = parseIntOrNull(end.tora), toda = parseIntOrNull(end.toda), asda = parseIntOrNull(end.asda);
        if (tora === null || toda === null || asda === null) {
          return { available: false, display: gaps.na(label + ": declared takeoff distances (TORA/TODA/ASDA) not fully published in NASR", "Verify via Chart Supplement / Airport Diagram") };
        }
        var limiting = Math.min(tora, toda, asda);
        var which = limiting === asda ? "ASDA" : limiting === toda ? "TODA" : "TORA";
        return { available: true, value: limiting, display: fmtNum(limiting) + " ft (" + which + ")" };
      }
      case "landingDistance": {
        var lda = parseIntOrNull(end.lda);
        if (lda === null) return { available: false, display: gaps.na(label + ": declared landing distance (LDA) not published in NASR", "Verify via Chart Supplement / Airport Diagram") };
        return { available: true, value: lda, display: fmtNum(lda) + " ft (LDA)" };
      }
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
  // Merged runway + criteria evaluation (Sections formerly 4 & 5, now one).
  // Evaluated per runway END, not per physical runway, because declared
  // distances are directional -- a displaced threshold or obstacle at one
  // end does not necessarily affect the opposite end.
  // -----------------------------------------------------------------------
  var RESULT_SEVERITY = { FAIL: 3, UNKNOWN: 2, CAUTION: 2, "NOT EVALUATED": 1, PASS: 0 };
  function worstResult(results) {
    var worst = "PASS";
    results.forEach(function (r) {
      if (RESULT_SEVERITY[r] > RESULT_SEVERITY[worst]) worst = r;
    });
    return worst;
  }

  function buildEndsAndCriteria(rec, criteriaData, gaps) {
    var endFacts = [];
    (rec.runways || []).forEach(function (rw) {
      (rw.ends || []).forEach(function (end) {
        endFacts.push({
          endId: end.id,
          runwayId: rw.id,
          width: rw.width,
          surface: rw.surfaceDesc || rw.surface || gaps.na("Runway " + rw.id + " surface not in NASR", "Verify via Chart Supplement"),
          wtBrg: weightBearingCompact(rw) || gaps.na("Runway " + rw.id + " weight bearing not in NASR", "Verify via Chart Supplement"),
          tora: end.tora || null,
          toda: end.toda || null,
          asda: end.asda || null,
          lda: end.lda || null,
          hasDisplacedThr: end.hasDisplacedThr,
          displacedThrLen: end.displacedThrLen,
          _rw: rw, _end: end,
        });
      });
    });

    var perAircraft = criteriaData.aircraft.map(function (aircraft) {
      var criteriaTables = [];
      endFacts.forEach(function (ef) {
        var rows = aircraft.criteria.map(function (def) {
          var actual = actualFor(def.id, rec, { rw: ef._rw, end: ef._end }, gaps);
          return evaluateCriterion(def, actual);
        });
        var verdict = worstResult(rows.map(function (r) { return r.result; }));
        criteriaTables.push({ endId: ef.endId, runwayId: ef.runwayId, rows: rows, verdict: verdict });
      });

      var qualifying = criteriaTables.filter(function (t) { return t.verdict !== "FAIL"; });
      var clean = qualifying.filter(function (t) { return t.verdict === "PASS"; });

      var determination;
      if (clean.length > 0) determination = "SUITABLE";
      else if (qualifying.length > 0) determination = "SUITABLE WITH LIMITATIONS";
      else determination = "NOT SUITABLE";

      var qualifyingIds = qualifying.map(function (t) { return t.endId; });
      var note = null;
      if (qualifying.length === 0) {
        note = aircraft.name + ": no runway end meets evaluated criteria (or none have been evaluated yet).";
      } else if (qualifying.length === 1) {
        note = aircraft.name + ": single qualifying runway end (" + qualifyingIds[0] + "). If it is out of service, no end on the field currently qualifies.";
      }

      return {
        aircraftId: aircraft.id,
        aircraftName: aircraft.name,
        determination: determination,
        qualifyingEnds: qualifyingIds,
        note: note,
        criteriaTables: criteriaTables,
      };
    });

    return { endFacts: endFacts, perAircraft: perAircraft };
  }

  // -----------------------------------------------------------------------
  // Services & FBO (Section 6) -- Fix 4: pilot-entered rows take priority
  // over (and are merged with) whatever's in fbo-data.js; anything still
  // blank prints CONFIRM.
  // -----------------------------------------------------------------------
  function buildFbo(icao, fboData, manualFbo, gaps) {
    var researched = (fboData && fboData[icao]) || [];
    var list = (manualFbo && manualFbo.length) ? manualFbo : researched;
    if (list.length === 0) return [];
    return list.map(function (fbo) {
      var name = fbo.name || gaps.confirm("Unnamed FBO record for " + icao, "Identify and confirm FBO name");
      return {
        name: name,
        phone: fbo.phone || gaps.confirm(name + " phone not confirmed", "Look up and confirm before dispatch"),
        fuel: fbo.fuel || gaps.confirm(name + " fuel availability not confirmed", "Call to confirm fuel"),
        services: fbo.services || (
          [fbo.gpu ? "GPU" : "", fbo.deice ? "De-ice" : "", fbo.hangar ? "Hangar" : "", fbo.hours ? "Hours: " + fbo.hours : ""]
            .filter(Boolean).join(", ") || fbo.sourceNote
        ) || gaps.confirm(name + " services not confirmed", "Call to confirm services"),
      };
    });
  }

  function airnavUrl(icao) {
    return "https://www.airnav.com/airport/" + icao;
  }

  // -----------------------------------------------------------------------
  // Disqualifying findings (Section 7) -- mechanical, derived only from
  // actual FAIL results already computed above. No new judgment calls.
  // -----------------------------------------------------------------------
  function buildDisqualifyingFindings(perAircraft) {
    var findings = [];
    var byEnd = {};
    perAircraft.forEach(function (a) {
      a.criteriaTables.forEach(function (t) {
        if (t.verdict !== "FAIL") return;
        var key = t.runwayId + "|" + t.endId;
        byEnd[key] = byEnd[key] || { runwayId: t.runwayId, endId: t.endId, aircraft: [] };
        byEnd[key].aircraft.push(a.aircraftName);
      });
    });
    Object.keys(byEnd).forEach(function (key) {
      var e = byEnd[key];
      findings.push("RWY " + e.runwayId + " (end " + e.endId + "): does not meet evaluated criteria for " + e.aircraft.join(" and ") + ".");
    });
    perAircraft.forEach(function (a) {
      if (a.determination === "NOT SUITABLE") {
        findings.push(a.aircraftName + " — airport: no runway end meets evaluated criteria for this aircraft.");
      }
    });
    return findings;
  }

  // -----------------------------------------------------------------------
  // Top-level build
  // -----------------------------------------------------------------------
  function buildReport(icao, airportData, criteriaData, fboData, extras) {
    extras = extras || {};
    var rec = (airportData || {})[icao];
    if (!rec) return { error: "No NASR record found for ICAO \"" + icao + "\" in the embedded dataset." };

    var gaps = makeGapTracker();

    var facility = buildFacility(rec, gaps);
    var communications = buildCommunications(rec);
    var endsAndCriteria = buildEndsAndCriteria(rec, criteriaData, gaps);
    var fbo = buildFbo(icao, fboData, extras.manualFbo, gaps);
    var disqualifying = buildDisqualifyingFindings(endsAndCriteria.perAircraft);

    return {
      icao: rec.icao,
      airportName: rec.name,
      city: rec.city,
      state: rec.state,
      verified: !!criteriaData.verified,
      criteriaNote: (criteriaData.meta && criteriaData.meta.note) || "",
      facility: facility,
      communications: communications,
      endFacts: endsAndCriteria.endFacts,
      determinations: endsAndCriteria.perAircraft,
      fbo: fbo,
      fboLookupUrl: airnavUrl(rec.icao),
      disqualifyingFindings: disqualifying,
      verificationItems: gaps.items,
      picComments: extras.picComments || "",
      picName: extras.picName || "",
      chiefPilotName: extras.chiefPilotName || "",
    };
  }

  return { buildReport: buildReport, NOT_AVAILABLE: NOT_AVAILABLE, CONFIRM: CONFIRM };
});
