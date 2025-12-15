(function () {
  "use strict";

  // ====== GUARD: prevent double run ======
  if (window.AV_TOOL && window.AV_TOOL.running) return;

  window.AV_TOOL = window.AV_TOOL || {};
  window.AV_TOOL.running = true;
  window.AV_TOOL.version = "1.0.0-esushrut-g5-full";
  window.AV_TOOL.panicHits = [];

  // ====== CONFIG ======
  const CONF = {
    autoDeselectOutOfRange: true,
    deselectZero: false,
    deselectNegative: false,

    highlightAbnormalBg: "rgba(254, 243, 199, 1)", // amber
    highlightPanicBg: "rgba(254, 226, 226, 1)",    // red tint

    waitSelector: "table",
    waitTimeoutMs: 12000
  };

  // ====== PANIC RULES (EDIT AS PER AIIMS POLICY) ======
  const PANIC_RULES = [
    { test: /glucose/i, unit: /mg\/dl/i, low: 40, high: 500 },
    { test: /sodium|na\b/i, unit: /mmol\/l|meq\/l/i, low: 120, high: 160 },
    { test: /potassium|k\b/i, unit: /mmol\/l|meq\/l/i, low: 2.5, high: 6.5 },
    { test: /calcium/i, unit: /mg\/dl/i, low: 6.0, high: 13.0 }
  ];

  // ====== REFERENCE RANGE OVERRIDES (COMMON ADULT STARTER TEMPLATE) ======
  // Used only when HIS "Reference Range" is missing / invalid / not parseable.
  const REF_OVERRIDES = [
    // GLUCOSE
    { test: /fasting.*glucose|fasting plasma glucose|\bfbs\b/i, unit: /mg\/dl/i, min: 70, max: 110 },
    { test: /random.*glucose|random plasma glucose|rbs/i, unit: /mg\/dl/i, min: 70, max: 140 },
    { test: /post.*prandial|postprandial|pp\b.*glucose/i, unit: /mg\/dl/i, min: 70, max: 140 },
    { test: /hba1c/i, unit: /%/i, min: 4.0, max: 5.6 },

    // ELECTROLYTES
    { test: /sodium|\bna\b/i, unit: /mmol\/l|meq\/l/i, min: 135, max: 145 },
    { test: /potassium|\bk\b/i, unit: /mmol\/l|meq\/l/i, min: 3.5, max: 5.1 },
    { test: /chloride|\bcl\b/i, unit: /mmol\/l|meq\/l/i, min: 98, max: 107 },
    { test: /bicarbonate|hco3|total co2/i, unit: /mmol\/l|meq\/l/i, min: 22, max: 28 },
    { test: /^calcium$|serum calcium|total calcium/i, unit: /mg\/dl/i, min: 8.6, max: 10.2 },
    { test: /^calcium$|serum calcium|total calcium/i, unit: /mmol\/l/i, min: 2.15, max: 2.55 },
    { test: /magnesium/i, unit: /mg\/dl/i, min: 1.7, max: 2.2 },
    { test: /phosphate|phosphorus/i, unit: /mg\/dl/i, min: 2.5, max: 4.5 },

    // RFT
    { test: /urea|blood urea/i, unit: /mg\/dl/i, min: 15, max: 40 },
    { test: /bun/i, unit: /mg\/dl/i, min: 7, max: 20 },
    { test: /creatinine/i, unit: /mg\/dl/i, min: 0.7, max: 1.3 },
    { test: /uric acid/i, unit: /mg\/dl/i, min: 3.5, max: 7.2 },

    // LFT
    { test: /bilirubin.*total|total bilirubin/i, unit: /mg\/dl/i, min: 0.3, max: 1.2 },
    { test: /bilirubin.*direct|direct bilirubin/i, unit: /mg\/dl/i, min: 0.0, max: 0.3 },
    { test: /bilirubin.*indirect|indirect bilirubin/i, unit: /mg\/dl/i, min: 0.2, max: 0.9 },
    { test: /\balt\b|sgpt/i, unit: /u\/l|iu\/l/i, min: 0, max: 40 },
    { test: /\bast\b|sgot/i, unit: /u\/l|iu\/l/i, min: 0, max: 40 },
    { test: /alkaline phosphatase|\balp\b/i, unit: /u\/l|iu\/l/i, min: 44, max: 147 },
    { test: /gamma.*gt|ggt/i, unit: /u\/l|iu\/l/i, min: 0, max: 60 },
    { test: /^total protein$|total proteins?/i, unit: /g\/dl/i, min: 6.4, max: 8.3 },
    { test: /albumin/i, unit: /g\/dl/i, min: 3.5, max: 5.0 },
    { test: /globulin/i, unit: /g\/dl/i, min: 2.0, max: 3.5 },
    { test: /a\/g|albumin.*globulin/i, unit: /.*/i, min: 1.0, max: 2.2 },

    // THYROID
    { test: /\btsh\b|thyroid stimulating hormone/i, unit: /u?iu\/ml|miu\/l/i, min: 0.4, max: 4.5 },
    { test: /free t4|\bft4\b/i, unit: /ng\/dl/i, min: 0.8, max: 1.8 },
    { test: /free t4|\bft4\b/i, unit: /pmol\/l/i, min: 10, max: 23 },
    { test: /free t3|\bft3\b/i, unit: /pg\/ml/i, min: 2.3, max: 4.2 },
    { test: /free t3|\bft3\b/i, unit: /pmol\/l/i, min: 3.5, max: 6.5 },
    { test: /^t4$|total t4/i, unit: /ug\/dl/i, min: 5.0, max: 12.0 },
    { test: /^t3$|total t3/i, unit: /ng\/dl/i, min: 80, max: 200 },

    // IRON STUDIES
    { test: /^iron$|serum iron/i, unit: /ug\/dl/i, min: 60, max: 170 },
    { test: /\btibc\b|total iron binding capacity/i, unit: /ug\/dl/i, min: 240, max: 450 },
    { test: /\buibc\b|unsaturated iron binding capacity/i, unit: /ug\/dl/i, min: 110, max: 370 },
    { test: /transferrin saturation|% ?saturation/i, unit: /%/i, min: 20, max: 50 },
    { test: /transferrin/i, unit: /mg\/dl/i, min: 200, max: 360 },

    // FERRITIN
    { test: /ferritin/i, unit: /ng\/ml|ug\/l/i, min: 30, max: 400 },

    // VITAMIN D
    { test: /vitamin d|25.*oh.*d|25-hydroxyvitamin d/i, unit: /ng\/ml/i, min: 20, max: 50 },

    // VITAMIN B12
    { test: /vitamin b12|cobalamin/i, unit: /pg\/ml/i, min: 200, max: 900 }
  ];

  // ====== UI (toast) ======
  function toast(msg, type = "info") {
    try {
      let t = document.getElementById("av-toast");
      if (!t) {
        t = document.createElement("div");
        t.id = "av-toast";
        t.style.cssText =
          "position:fixed;top:16px;left:50%;transform:translateX(-50%);z-index:2147483647;" +
          "padding:10px 16px;border-radius:10px;font:600 14px/1.2 system-ui,Segoe UI,Arial;" +
          "box-shadow:0 10px 25px rgba(0,0,0,.25);opacity:0;transition:opacity .2s;" +
          "pointer-events:none;";
        document.body.appendChild(t);
      }
      t.style.background =
        type === "success" ? "#10b981" : type === "error" ? "#ef4444" : "#3b82f6";
      t.style.color = "#fff";
      t.textContent = msg;
      t.style.opacity = "1";
      setTimeout(() => (t.style.opacity = "0"), 2200);
    } catch (e) {
      console.log(msg);
    }
  }

  // ====== Parsing helpers ======
  function parseNum(s) {
    if (s == null) return null;
    s = ("" + s).replace(/\u00A0/g, " ").replace(/,/g, " ").trim();
    const m = s.match(/-?\d+(\.\d+)?/);
    return m ? Number(m[0]) : null;
  }

  function parseRange(s) {
    if (!s) return null;
    s = ("" + s).replace(/\u00A0/g, " ").replace(/,/g, " ").trim();
    const m = s.match(/(-?\d+(\.\d+)?)\s*(?:-|–|to)\s*(-?\d+(\.\d+)?)/i);
    if (!m) return null;
    const a = Number(m[1]), b = Number(m[3]);
    if (!Number.isFinite(a) || !Number.isFinite(b)) return null;
    return { min: Math.min(a, b), max: Math.max(a, b) };
  }

  function guessUnit(text) {
    const t = (text || "").toLowerCase();
    const m = t.match(/mg\/dl|mmol\/l|meq\/l|u\/l|iu\/l|ng\/ml|pg\/ml|ug\/dl|ug\/l|g\/dl|miu\/l|uiu\/ml|pmol\/l|ng\/dl|%/i);
    return m ? m[0] : "";
  }

  function getOverrideRange(testName, unitText) {
    const unit = guessUnit(unitText);
    for (const r of REF_OVERRIDES) {
      if (r.test.test(testName) && (!r.unit || r.unit.test(unit || unitText))) {
        return { min: r.min, max: r.max, source: "override", unit };
      }
    }
    return null;
  }

  // ====== Checkbox helpers ======
  function isChecked(cb) {
    return !!(cb && (cb.checked || cb.getAttribute("aria-checked") === "true"));
  }

  function forceUncheck(cb) {
    if (!cb || cb.disabled) return;
    try {
      cb.checked = false;
      cb.setAttribute("aria-checked", "false");
      cb.dispatchEvent(new Event("input", { bubbles: true }));
      cb.dispatchEvent(new Event("change", { bubbles: true }));
    } catch (e) {}
    try {
      if (isChecked(cb)) cb.click();
    } catch (e) {}
    try {
      cb.checked = false;
      cb.setAttribute("aria-checked", "false");
    } catch (e) {}
  }

  // ====== Panic helpers ======
  function matchPanicRule(testName, unitText) {
    for (const r of PANIC_RULES) {
      if (r.test.test(testName) && (!r.unit || r.unit.test(unitText || ""))) return r;
    }
    return null;
  }

  function isPanic(value, rule) {
    if (value == null || !rule) return false;
    if (rule.low != null && value < rule.low) return true;
    if (rule.high != null && value > rule.high) return true;
    return false;
  }

  // ====== Lock Save & Validate (eSushrut G5) ======
  function lockESushrutSaveValidate(lock = true) {
    const buttons = [...document.querySelectorAll("button")];
    buttons.forEach((btn) => {
      const txt = (btn.innerText || "").toLowerCase().trim();
      const isSaveValidate = txt.includes("save") && txt.includes("validate");
      if (!isSaveValidate) return;

      if (lock) {
        if (btn.dataset.avPrevDisabled == null) {
          btn.dataset.avPrevDisabled = btn.disabled ? "1" : "0";
        }
        btn.disabled = true;
        btn.style.opacity = "0.5";
        btn.style.cursor = "not-allowed";
      } else {
        if (btn.dataset.avPrevDisabled === "0") {
          btn.disabled = false;
          btn.style.opacity = "";
          btn.style.cursor = "";
        }
      }
    });
  }

  window.AV_TOOL.acknowledgePanic = function () {
    lockESushrutSaveValidate(false);
    window.AV_TOOL.panicHits = [];
    toast("Panic acknowledged. Save & Validate unlocked.", "success");
  };

  // ====== Table mapping for eSushrut G5 ======
  function detectTableMapping(table) {
    const headerRow = table.querySelector("tr");
    if (!headerRow) return null;

    const headers = [...headerRow.querySelectorAll("th,td")].map((h) =>
      (h.innerText || "").toLowerCase().trim()
    );

    const nameIdx = headers.findIndex((h) =>
      h.includes("test param name") || h.includes("param name") || h.includes("test name")
    );
    const valIdx = headers.findIndex((h) =>
      h.includes("test param value") || h.includes("param value") || h.includes("value") || h.includes("result")
    );
    const rngIdx = headers.findIndex((h) =>
      h.includes("reference range") || h.includes("ref") || h.includes("range") || h.includes("normal")
    );

    if (valIdx !== -1 && rngIdx !== -1) return { nameIdx, valIdx, rngIdx };
    return null;
  }

  // ====== Core runner ======
  function run() {
    const stats = { matched: 0, deselected: 0, panic: 0, tables: 0 };

    window.AV_TOOL.panicHits = [];

    const tables = [...document.querySelectorAll("table")];
    tables.forEach((table) => {
      const map = detectTableMapping(table);
      if (!map) return;

      stats.tables++;

      const rows = [...table.querySelectorAll("tr")];
      rows.forEach((row, idx) => {
        if (idx === 0) return; // header

        const cb = row.querySelector('input[type="checkbox"]');
        if (!cb || !isChecked(cb)) return;

        const cells = [...row.querySelectorAll("td,th")];
        if (!cells[map.valIdx] || !cells[map.rngIdx]) return;

        const testName =
          map.nameIdx !== -1 && cells[map.nameIdx]
            ? (cells[map.nameIdx].innerText || "").trim()
            : (cells[0]?.innerText || "").trim();

        const valText = (cells[map.valIdx].innerText || "").trim();
        const rngText = (cells[map.rngIdx].innerText || "").trim();

        const value = parseNum(valText);

        // Range selection: HIS first, else override
        let range = parseRange(rngText);
        let rangeSource = "his";
        if (!range) {
          const ov = getOverrideRange(testName, rngText);
          if (ov) {
            range = { min: ov.min, max: ov.max };
            rangeSource = "override";
          }
        }

        if (value == null || !range) return;

        stats.matched++;

        const unitGuess = guessUnit(rngText) || rngText;
        const panicRule = matchPanicRule(testName, unitGuess);

        // PANIC: highlight + lock; do NOT deselect
        if (isPanic(value, panicRule)) {
          stats.panic++;

          window.AV_TOOL.panicHits.push({
            test: testName,
            value: value,
            unit: guessUnit(rngText) || ""
          });

          row.style.backgroundColor = CONF.highlightPanicBg;
          try {
            cells[map.valIdx].style.border = "2px solid #dc2626";
            cells[map.valIdx].style.fontWeight = "800";
          } catch (e) {}
          return;
        }

        const outOfRange = value < range.min || value > range.max;

        const shouldDeselect =
          (CONF.autoDeselectOutOfRange && outOfRange) ||
          (CONF.deselectZero && value === 0) ||
          (CONF.deselectNegative && value < 0);

        if (shouldDeselect) {
          forceUncheck(cb);
          stats.deselected++;

          row.style.backgroundColor = CONF.highlightAbnormalBg;
          try {
            cells[map.valIdx].style.border = "2px solid #ef4444";
            cells[map.valIdx].style.fontWeight = "700";
            if (rangeSource === "override") row.title = "Reference range used: override (AIIMS template)";
          } catch (e) {}
        }
      });
    });

    if (window.AV_TOOL.panicHits.length) {
      lockESushrutSaveValidate(true);

      const msg = window.AV_TOOL.panicHits
        .map((p) => `• ${p.test}: ${p.value} ${p.unit}`.trim())
        .join("\n");

      alert(
        "⚠️ PANIC VALUES DETECTED ⚠️\n\n" +
          msg +
          "\n\nSave & Validate is LOCKED.\n" +
          "Review the report, then run in Console:\n" +
          "AV_TOOL.acknowledgePanic()"
      );
    }

    toast(
      `Done: tables ${stats.tables}, matched ${stats.matched}, deselected ${stats.deselected}, panic ${stats.panic}`,
      stats.panic ? "error" : "success"
    );
  }

  // ====== Wait-for-page helper (eSushrut loads dynamically) ======
  function waitFor(selector, timeoutMs) {
    const start = Date.now();
    const timer = setInterval(() => {
      if (document.querySelector(selector)) {
        clearInterval(timer);
        run();
      } else if (Date.now() - start > timeoutMs) {
        clearInterval(timer);
        toast("AutoValidator: table not detected (timeout).", "error");
      }
    }, 250);
  }

  toast(`AutoValidator v${window.AV_TOOL.version}: scanning…`, "info");
  waitFor(CONF.waitSelector, CONF.waitTimeoutMs);
})();
