// Pure computation layer for the airport screening report.
//
// This is a FIRST-PASS SCREENING tool, not a performance-suitability
// determination -- the PACCAR FOM confirms no static minimum runway length
// exists, so length/width/weight performance is deferred to a real
// ForeFlight performance calculation for the day's weight/weather. What
// this module screens for is the FOM's own gating questions: 2-year
// recency (ARA trigger, §2.7.4), instrument approach availability (drives
// alternate minimums, §4.1.3.3), night capability and remote/unmonitored
// field risk (§6.1.6.7), night circling (§6.1.6.2), and the one truly
// static runway fact -- hard surface + weight-bearing class.
//
// Takes NASR airport data + criteria.json + the FBO directory (plus
// pilot-entered extras: manual FBO rows, PIC comments, sign-off names, and
// the three FOM screening answers) and produces a plain-JS "report model"
// -- no DOM, no docx library -- so it can run identically in the browser
// (index.html) and under Node (verify-report.js).
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
//   - This tool never says "check with the DOA" and never claims to be
//     ForeFlight. It screens; marginal cases route to pilot + chief pilot
//     (FOM §2.7.4), which is the FOM's own process, not a cop-out.
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

  var BUCKET = {
    DISQUALIFIER: "CLEAR DISQUALIFIER",
    REVIEW: "REVIEW WITH CHIEF PILOT",
    OK: "SCREENS OK — PERFORMANCE CALC REQUIRED",
  };

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

  var APPROACH_LABEL = {
    none: "None (VFR only)",
    nonPrecision: "Non-precision only",
    precisionOrApv: "Precision or APV",
    unknown: "Unknown — verify from chart",
  };

  // Communications whitelist. NASR's FRQ.csv lists an airport's TRACON-
  // served named arrival/departure procedure frequencies (one row per
  // STAR/DP/RNAV transition) right alongside its own tower/ground/CTAF/
  // ATIS rows. Those procedure rows are real data but they are FMS routing
  // noise for a screening report -- a busy Class-B-served field can have
  // 30+ of them burying the one tower frequency a crew needs.
  var NAMED_PROCEDURE_RE = /\b(STAR|DP|RNAV)\b/;
  function isOperationalFrequency(f) {
    var use = (f.use || "").toUpperCase().trim();
    if (!use) return false;
    if (NAMED_PROCEDURE_RE.test(use)) return false; // named arrival/departure procedure -- drop regardless of category
    var freqNum = parseFloat(f.freq);
    if (isFinite(freqNum) && freqNum >= 225) return false; // UHF military duplicate of a civil VHF frequency
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
  function parseIntOrNull(v) {
    if (v === undefined || v === null || v === "") return null;
    var n = parseInt(v, 10);
    return isFinite(n) ? n : null;
  }

  // Verification-item accumulator. Every "gap" the tool renders (NOT
  // AVAILABLE from NASR, CONFIRM on an unverified FBO fact, or a marginal
  // screening condition) is logged here exactly once (by item text) so
  // Section 8 can be auto-generated straight from what the rest of the
  // report actually rendered.
  function makeGapTracker() {
    var seen = {};
    var items = [];
    function log(item, action) {
      if (!seen[item]) { seen[item] = true; items.push({ item: item, action: action }); }
    }
    return {
      na: function (item, action) { log(item, action); return NOT_AVAILABLE; },
      confirm: function (item, action) { log(item, action); return CONFIRM; },
      log: log,
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
    return { letter: letter, display: "Index " + letter + " (FAR 139 Class " + parts[0] + ")" };
  }

  // -----------------------------------------------------------------------
  // Communications (Section 3) -- operational set only, de-duped.
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
  // criterion has never been reviewed against a real company number, so no
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
      var margin = actual.value - def.threshold;
      var sign = margin >= 0 ? "+" : "-";
      row.marginDisplay = sign + fmtNum(Math.abs(margin)) + (def.unit ? " " + def.unit : "");
      row.result = margin >= 0 ? "PASS" : "FAIL";
      if (row.result === "PASS" && def.cautionBand && margin < def.cautionBand) row.result = "CAUTION";
    } else if (def.compare === "categorical") {
      row.result = def.allowed.indexOf(actual.value) !== -1 ? "PASS" : "FAIL";
    }
    return row;
  }

  function requiredDisplayFor(def) {
    if (def.compare === "categorical") return def.requiredLabel || def.allowed.join(", ");
    var unit = def.unit ? " " + def.unit : "";
    return fmtNum(def.threshold) + unit;
  }

  function actualForRunway(criterionId, rw, gaps) {
    switch (criterionId) {
      case "surface": {
        if (!rw.surface) return { available: false, display: gaps.na("Runway " + rw.id + " surface not in NASR", "Verify via Chart Supplement") };
        var baseToken = rw.surface.split(/[-/]/)[0];
        return { available: true, value: baseToken, display: rw.surfaceDesc || rw.surface };
      }
      case "wtBearingDW": {
        var dw = parseIntOrNull(rw.grossWtDW);
        if (dw === null) return { available: false, display: gaps.na("Runway " + rw.id + " dual-wheel weight bearing not in NASR", "Verify PCN/gross weight via Chart Supplement") };
        return { available: true, value: dw * 1000, display: fmtNum(dw * 1000) + " lb" };
      }
      default:
        return { available: false, display: NOT_AVAILABLE };
    }
  }

  var RESULT_SEVERITY = { FAIL: 3, UNKNOWN: 2, CAUTION: 2, "NOT EVALUATED": 1, PASS: 0 };
  function worstResult(results) {
    var worst = "PASS";
    results.forEach(function (r) { if (RESULT_SEVERITY[r] > RESULT_SEVERITY[worst]) worst = r; });
    return worst;
  }

  // -----------------------------------------------------------------------
  // Runway facts (Section 4) -- declared distances are now a displayed
  // fact for the performance handoff to ForeFlight, never a pass/fail
  // gate, since length screening has no static company minimum. The only
  // per-runway screening left is hard-surface + weight-bearing class.
  // -----------------------------------------------------------------------
  function buildRunwayFactsAndScreening(rec, criteriaData, gaps) {
    var endFacts = [];
    (rec.runways || []).forEach(function (rw) {
      (rw.ends || []).forEach(function (end) {
        endFacts.push({
          endId: end.id, runwayId: rw.id,
          tora: end.tora || null, toda: end.toda || null, asda: end.asda || null, lda: end.lda || null,
          hasDisplacedThr: end.hasDisplacedThr, displacedThrLen: end.displacedThrLen,
        });
      });
    });

    var runwayFacts = (rec.runways || []).map(function (rw) {
      return {
        id: rw.id,
        width: rw.width,
        surface: rw.surfaceDesc || rw.surface || gaps.na("Runway " + rw.id + " surface not in NASR", "Verify via Chart Supplement"),
        wtBrg: weightBearingCompact(rw) || gaps.na("Runway " + rw.id + " weight bearing not in NASR", "Verify via Chart Supplement"),
      };
    });

    var perAircraft = criteriaData.aircraft.map(function (aircraft) {
      var runwayScreens = (rec.runways || []).map(function (rw) {
        var rows = aircraft.criteria.map(function (def) {
          var actual = actualForRunway(def.id, rw, gaps);
          return evaluateCriterion(def, actual);
        });
        return { runwayId: rw.id, rows: rows, verdict: worstResult(rows.map(function (r) { return r.result; })) };
      });
      return { aircraftId: aircraft.id, aircraftName: aircraft.name, runwayScreens: runwayScreens };
    });

    return { endFacts: endFacts, runwayFacts: runwayFacts, perAircraft: perAircraft };
  }

  // -----------------------------------------------------------------------
  // Instrument approach classification -- auto-derived from the ground-
  // based NAVAID field where possible, always overridable. NASR's ILS_TYPE
  // does NOT cover RNAV(GPS)-only approaches (separate FAA TPP data), so a
  // blank there can only mean "unknown", never "none" -- absence is not
  // evidence of absence.
  // -----------------------------------------------------------------------
  function deriveApproachType(rec) {
    var hasILS = false, hasOtherNavaid = false;
    (rec.runways || []).forEach(function (rw) {
      (rw.ends || []).forEach(function (end) {
        var t = (end.ilsType || "").toUpperCase();
        if (!t) return;
        if (t.indexOf("ILS") !== -1) hasILS = true;
        else hasOtherNavaid = true;
      });
    });
    if (hasILS) return "precisionOrApv";
    if (hasOtherNavaid) return "nonPrecision";
    return null; // cannot derive -- NOT the same as "none"
  }

  function buildNightOpsAndRecency(rec, communications, criteriaData, gaps, extras) {
    // 2-year recency / ARA trigger (FOM §2.7.4).
    var recencyAnswer = extras.recentOperation || "unknown";
    var recencyNote;
    if (recencyAnswer === "yes") {
      recencyNote = "Crew has operated into this airport within the last 2 years. An Airport Risk Assessment (ARA) may already exist in Comply365 — check before creating a new one (FOM §2.7.4).";
    } else if (recencyAnswer === "no") {
      recencyNote = "Neither crew member has operated into this airport within the last 2 years. A new Airport Risk Assessment (ARA) may be required (FOM §2.7.4).";
      gaps.log("2-year recency: neither crew member has recent experience at " + rec.icao, "Confirm ARA requirement with Chief Pilot per FOM §2.7.4");
    } else {
      recencyNote = "2-year recency not confirmed. Ask: has either crew member operated into this airport within the last 2 years? If so, an ARA may already exist in Comply365 (FOM §2.7.4).";
      gaps.log("2-year recency not confirmed for " + rec.icao, "Ask the crew and update before dispatch (FOM §2.7.4 ARA trigger)");
    }

    // Instrument approach classification (drives alternate minimums, §4.1.3.3).
    var auto = deriveApproachType(rec);
    var override = extras.instrumentApproachOverride || "";
    var effective = override || auto || "unknown";
    var manualEntryRequired = !auto;
    if (effective === "unknown") {
      gaps.log("Instrument approach availability not confirmed for " + rec.icao, "Verify via current IAP charts / ForeFlight before dispatch");
    }
    var altMin = null;
    if (effective === "precisionOrApv") altMin = criteriaData.fom.alternateMinimums.precisionOrApv;
    else if (effective === "nonPrecision") altMin = criteriaData.fom.alternateMinimums.nonPrecision;
    var vfrOnlyFlag = effective === "none";

    // Night capability (§6.1.6.7): edge lighting, lighted wind indicator, obstruction lighting.
    var edgeLightingPresent = (rec.runways || []).some(function (rw) { return rw.lighting && rw.lighting !== "NONE"; });
    var windIndicatorLighted = rec.windIndicatorFlag === "Y-L";
    var obstructionLightingPresent = (rec.runways || []).some(function (rw) {
      return (rw.ends || []).some(function (e) { return /L/.test(e.obstnMrkdCode || ""); });
    });

    // Remote / unmonitored (§6.1.6.7): uncontrolled + no on-field weather reporting.
    var hasOnFieldWeather = communications.some(function (c) { return c.facilityType === "ASOS_AWOS"; });
    var remoteUnmonitored = !rec.tower && !hasOnFieldWeather;
    var winterNightOps = extras.winterNightOps || "unknown";
    var winterNightHardFlag = remoteUnmonitored && winterNightOps === "yes";
    if (remoteUnmonitored && !winterNightHardFlag) {
      gaps.log(rec.icao + " is uncontrolled with no on-field weather reporting", "Review before any night operation, especially in winter (FOM §6.1.6.7)");
    }

    return {
      recency: { answer: recencyAnswer, note: recencyNote },
      instrumentApproach: {
        auto: auto, autoLabel: auto ? APPROACH_LABEL[auto] : null,
        override: override || null, effective: effective, effectiveLabel: APPROACH_LABEL[effective],
        manualEntryRequired: manualEntryRequired, overridden: !!override && override !== auto,
      },
      alternateMinimums: altMin,
      vfrOnlyFlag: vfrOnlyFlag,
      nightCapability: { edgeLightingPresent: edgeLightingPresent, windIndicatorLighted: windIndicatorLighted, obstructionLightingPresent: obstructionLightingPresent },
      nightCirclingNote: "Night circling approaches are prohibited without Chief Pilot / DOA approval (FOM §6.1.6.2).",
      remoteUnmonitored: remoteUnmonitored,
      winterNightOps: winterNightOps,
      winterNightHardFlag: winterNightHardFlag,
    };
  }

  // -----------------------------------------------------------------------
  // Determination (Section 1) -- three screening buckets. Never SUITABLE /
  // NOT SUITABLE: this tool screens, it does not determine performance
  // suitability, and marginal cases route to pilot + chief pilot (FOM
  // §2.7.4), not a silent pass/fail.
  // -----------------------------------------------------------------------
  function computeDetermination(runwayScreens, nightOps) {
    var usable = runwayScreens.filter(function (r) {
      var s = r.rows.filter(function (x) { return x.id === "surface"; })[0];
      var w = r.rows.filter(function (x) { return x.id === "wtBearingDW"; })[0];
      return s.result !== "FAIL" && w.result !== "FAIL";
    });

    if (nightOps.winterNightHardFlag) {
      return {
        bucket: BUCKET.DISQUALIFIER,
        reasons: ["Uncontrolled field with no on-field weather reporting, and this is planned as a winter-season night operation (FOM §6.1.6.7)."],
      };
    }
    if (usable.length === 0) {
      return { bucket: BUCKET.DISQUALIFIER, reasons: ["No runway at this airport has a hard (paved) surface / meets the weight-bearing screen."] };
    }

    var reasons = [];
    if (nightOps.instrumentApproach.effective === "unknown") {
      reasons.push("Instrument approach availability not confirmed — alternate minimums cannot be stated until verified.");
    }
    if (nightOps.remoteUnmonitored && !nightOps.winterNightHardFlag) {
      reasons.push("Uncontrolled field with no on-field weather reporting — review before any night operation, especially in winter (FOM §6.1.6.7).");
    }
    var marginalWeight = usable.some(function (r) {
      var w = r.rows.filter(function (x) { return x.id === "wtBearingDW"; })[0];
      return w.result === "CAUTION" || w.result === "UNKNOWN";
    });
    if (marginalWeight) reasons.push("Weight-bearing data marginal or unavailable for at least one runway.");

    if (reasons.length) return { bucket: BUCKET.REVIEW, reasons: reasons };
    return { bucket: BUCKET.OK, reasons: [] };
  }

  // -----------------------------------------------------------------------
  // Services & FBO (Section 6) -- pilot-entered rows take priority over
  // (and are merged with) whatever's in fbo-data.js; anything still blank
  // prints CONFIRM.
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
  // Top-level build
  // -----------------------------------------------------------------------
  function buildReport(icao, airportData, criteriaData, fboData, extras) {
    extras = extras || {};
    var rec = (airportData || {})[icao];
    if (!rec) return { error: "No NASR record found for ICAO \"" + icao + "\" in the embedded dataset." };

    var gaps = makeGapTracker();

    var facility = buildFacility(rec, gaps);
    var communications = buildCommunications(rec);
    var runwayData = buildRunwayFactsAndScreening(rec, criteriaData, gaps);
    var nightOps = buildNightOpsAndRecency(rec, communications, criteriaData, gaps, extras);
    var fbo = buildFbo(icao, fboData, extras.manualFbo, gaps);

    var determinations = runwayData.perAircraft.map(function (a) {
      var det = computeDetermination(a.runwayScreens, nightOps);
      return { aircraftId: a.aircraftId, aircraftName: a.aircraftName, runwayScreens: a.runwayScreens, bucket: det.bucket, reasons: det.reasons };
    });

    var disqualifyingFindings = [];
    determinations.forEach(function (d) {
      if (d.bucket === BUCKET.DISQUALIFIER) {
        d.reasons.forEach(function (r) {
          var text = d.aircraftName + ": " + r;
          if (disqualifyingFindings.indexOf(text) === -1) disqualifyingFindings.push(text);
        });
      }
    });

    return {
      icao: rec.icao,
      airportName: rec.name,
      city: rec.city,
      state: rec.state,
      verified: !!criteriaData.verified,
      criteriaNote: (criteriaData.meta && criteriaData.meta.note) || "",
      facility: facility,
      communications: communications,
      endFacts: runwayData.endFacts,
      runwayFacts: runwayData.runwayFacts,
      determinations: determinations,
      nightOps: nightOps,
      fbo: fbo,
      fboLookupUrl: airnavUrl(rec.icao),
      disqualifyingFindings: disqualifyingFindings,
      verificationItems: gaps.items,
      picComments: extras.picComments || "",
      picName: extras.picName || "",
      chiefPilotName: extras.chiefPilotName || "",
    };
  }

  return {
    buildReport: buildReport,
    deriveApproachType: deriveApproachType,
    BUCKET: BUCKET,
    APPROACH_LABEL: APPROACH_LABEL,
    NOT_AVAILABLE: NOT_AVAILABLE,
    CONFIRM: CONFIRM,
  };
});
