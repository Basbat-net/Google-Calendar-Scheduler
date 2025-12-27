/************ CONFIG ************/
const CFG = {
  sheetName: "EstudioBien",

  // ===== CONTINUA =====
  continuaRangeA1: "E11:I16",          // Concepto | % | Nota mínima | Necesario | Obtenido
  continuaRecuperableColA1: "J11:J16", // Recuperable? (checkbox) alineado con continuaRangeA1

  // ===== ORDINARIA =====
  ordinariaRangeA1: "E18:I19",         // (2 filas) Laboratorio | Examen  -> Concepto | % | Nota mínima | Necesario | Obtenido

  // Columnas dentro del rango (1-based)
  colPct: 2,         // %
  colMin: 3,         // Nota mínima
  colNecesario: 4,   // Necesario (se escribe aquí)
  colObtenido: 5,    // Obtenido

  objective: 5.0,
  maxGrade: 10.0,
  step: 0.25,
  blackBg: "#000000",
  whiteBg: "#ffffff",
  impossibleText: "COOKED"
};

function recalcularNecesariosContinua() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.sheetName);
  if (!sh) throw new Error(`No se encontró la hoja "${CFG.sheetName}". Revisa CFG.sheetName.`);

  const objective = CFG.objective;
  const maxGrade = CFG.maxGrade;
  const step = CFG.step;

  /******************************************************************
   * 1) CONTINUA: calcular "Necesario" en CFG.continuaRangeA1
   ******************************************************************/
  const rg = sh.getRange(CFG.continuaRangeA1);
  const vals = rg.getValues();
  const n = vals.length;

  // Columnas (0-based) dentro del rango E..I
  const idxConcepto = 0;
  const idxPct = CFG.colPct - 1;
  const idxMin = CFG.colMin - 1;
  const idxNecesario = CFG.colNecesario - 1; // no se lee, solo mapping
  const idxObtenido = CFG.colObtenido - 1;

  const w = [];
  const minG = [];
  const got = [];
  const gotRaw = [];
  const concepto = [];

  for (let i = 0; i < n; i++) {
    concepto[i] = String(vals[i][idxConcepto] ?? "").trim();

    const pct = parseNumber(vals[i][idxPct]);
    w[i] = pct !== null ? pct / 100 : 0;

    minG[i] = parseNumber(vals[i][idxMin]) ?? 0;

    gotRaw[i] = vals[i][idxObtenido];
    got[i] = isBlank(gotRaw[i]) ? null : (parseNumber(gotRaw[i]) ?? 0);
  }

  // Salida (Necesario + fondos)
  const outValues = Array.from({ length: n }, () => [""]);
  const outBgs = Array.from({ length: n }, () => [CFG.whiteBg]);

  // Obtenidos → Necesario vacío + negro
  for (let i = 0; i < n; i++) {
    if (got[i] !== null) {
      outValues[i][0] = "";
      outBgs[i][0] = CFG.blackBg;
    }
  }

  // Pendientes con peso (>0) y sin nota
  const pending = [];
  for (let i = 0; i < n; i++) {
    if (w[i] > 0 && got[i] === null) pending.push(i);
  }

  // Aportación conseguida
  let achieved = 0;
  for (let i = 0; i < n; i++) {
    if (w[i] > 0 && got[i] !== null) achieved += w[i] * got[i];
  }

  // Filas con 0% y sin nota -> mostrar mínimo como "Necesario" (si existe)
  for (let i = 0; i < n; i++) {
    if (w[i] === 0 && got[i] === null) {
      if (minG[i] !== null && minG[i] > 0) {
        outValues[i][0] = roundUpToStep(minG[i], step);
        outBgs[i][0] = CFG.whiteBg;
      } else {
        outValues[i][0] = "";
        outBgs[i][0] = CFG.whiteBg;
      }
    }
  }

  // Resolver pendientes con peso
  if (pending.length === 0) {
    writeNecesario(rg, outValues, outBgs);
  } else {
    let minContribution = 0;
    let maxContribution = 0;
    for (const i of pending) {
      minContribution += w[i] * minG[i];
      maxContribution += w[i] * maxGrade;
    }

    if (achieved + maxContribution < objective - 1e-12) {
      for (const i of pending) outValues[i][0] = CFG.impossibleText;
      writeNecesario(rg, outValues, outBgs);
    } else if (achieved + minContribution >= objective - 1e-12) {
      for (const i of pending) outValues[i][0] = roundUpToStep(minG[i], step);
      writeNecesario(rg, outValues, outBgs);
    } else {
      const minOfMins = Math.min(...pending.map(i => minG[i]));

      function totalAtLevel(L) {
        let total = achieved;
        for (const i of pending) {
          const gi = clamp(L, minG[i], maxGrade);
          total += w[i] * gi;
        }
        return total;
      }

      let lo = minOfMins;
      let hi = maxGrade;
      for (let iter = 0; iter < 60; iter++) {
        const mid = (lo + hi) / 2;
        if (totalAtLevel(mid) >= objective) hi = mid;
        else lo = mid;
      }
      const L = hi;

      const grade = new Array(n).fill(null);
      for (const i of pending) {
        const gi = clamp(L, minG[i], maxGrade);
        grade[i] = roundUpToStep(gi, step);
      }

      function currentTotal() {
        let total = achieved;
        for (const i of pending) total += w[i] * grade[i];
        return total;
      }

      let total = currentTotal();

      while (true) {
        const candidates = [];
        for (const i of pending) {
          const newVal = grade[i] - step;
          if (newVal + 1e-12 < minG[i]) continue;
          if (total - w[i] * step + 1e-12 < objective) continue;
          candidates.push(i);
        }

        if (candidates.length === 0) break;

        candidates.sort((a, b) => {
          if (grade[b] !== grade[a]) return grade[b] - grade[a];
          return w[a] - w[b];
        });

        const k = candidates[0];
        grade[k] -= step;
        grade[k] = Math.round(grade[k] / step) * step;
        total -= w[k] * step;
      }

      for (const i of pending) {
        outValues[i][0] = grade[i];
        outBgs[i][0] = CFG.whiteBg;
      }

      writeNecesario(rg, outValues, outBgs);
    }
  }

  /******************************************************************
   * 2) ORDINARIA: calcular "Necesario" en CFG.ordinariaRangeA1
   *
   * Fixes pedidos:
   * - El LABORATORIO de ordinaria se basa en el OBTENIDO de CONTINUA (si existe).
   *   Si en continua hay un 7 y el mínimo es 5 -> ordinaria lab "Necesario" se deja vacío y en negro.
   * - Evitar doble conteo: el laboratorio de continua NO se suma dentro de "no recuperables",
   *   porque ya se trata aparte en ordinaria.
   ******************************************************************/
  const rgRec = sh.getRange(CFG.continuaRecuperableColA1);
  const recVals = rgRec.getValues(); // [[true/false], ...]

  // El laboratorio de continua lo tomamos como la PRIMERA fila del rango (tal como tu tabla)
  const contLabPct = parseNumber(vals[0][idxPct]) ?? 0;
  const contLabW = contLabPct / 100;
  const contLabMin = parseNumber(vals[0][idxMin]) ?? 0;
  const contLabGot = got[0]; // ya parseado arriba (null si vacío)

  // Suma ponderada de NO recuperables de continua EXCLUYENDO laboratorio (fila 0)
  let nonRecoverableAchieved = 0;
  for (let i = 1; i < n; i++) {
    const pct = parseNumber(vals[i][idxPct]);
    const wi = pct !== null ? pct / 100 : 0;
    const isRec = !!(recVals[i] && recVals[i][0] === true);

    if (wi > 0 && !isRec) {
      const gotNR = got[i];
      if (gotNR !== null) nonRecoverableAchieved += wi * gotNR;
    }
  }

  // Leer ORDINARIA
  const rgOrd = sh.getRange(CFG.ordinariaRangeA1);
  const ordVals = rgOrd.getValues();
  const ordN = ordVals.length;
  if (ordN < 2) throw new Error(`CFG.ordinariaRangeA1 debe incluir al menos 2 filas (Laboratorio y Examen).`);

  // Índices 0-based dentro ORDINARIA
  const oPct = CFG.colPct - 1;
  const oMin = CFG.colMin - 1;
  const oObt = CFG.colObtenido - 1;

  // ORDINARIA fila 0: Laboratorio (pero usamos el obtenido de CONTINUA)
  const labPct = parseNumber(ordVals[0][oPct]) ?? 0;
  const labW = labPct / 100;
  const labMin = parseNumber(ordVals[0][oMin]) ?? 0;

  // ORDINARIA fila 1: Examen
  const examPct = parseNumber(ordVals[1][oPct]) ?? 0; // viene de tu SUMIF externo
  const examW = examPct / 100;
  const examMin = parseNumber(ordVals[1][oMin]) ?? 0;

  const outOrdValues = Array.from({ length: ordN }, () => [""]);
  const outOrdBgs = Array.from({ length: ordN }, () => [CFG.whiteBg]);

  // Laboratorio ordinaria: si continua lab tiene nota >= mínimo, se usa esa nota y se marca necesario en negro
  // Si no, se exige mínimo.
  let labUsedForCalc;
  const labSourceGot = contLabGot; // puede ser null
  const labSourceMin = contLabMin; // mínimo en continua (normalmente igual al de ordinaria, pero no asumimos)

  // El mínimo efectivo para lab lo tomamos como el mayor de ambos mínimos por seguridad.
  const labEffectiveMin = Math.max(labMin ?? 0, labSourceMin ?? 0);

  if (labSourceGot !== null && labSourceGot + 1e-12 >= labEffectiveMin) {
    outOrdValues[0][0] = "";
    outOrdBgs[0][0] = CFG.blackBg;
    labUsedForCalc = labSourceGot;
  } else {
    outOrdValues[0][0] = roundUpToStep(labEffectiveMin, step);
    outOrdBgs[0][0] = CFG.whiteBg;
    labUsedForCalc = labEffectiveMin;
  }

  // Examen final: calcular necesario
  const achievedSoFar = nonRecoverableAchieved + labW * labUsedForCalc;

  if (examW <= 0) {
    if (achievedSoFar + 1e-12 < objective) outOrdValues[1][0] = CFG.impossibleText;
    else outOrdValues[1][0] = "";
    writeNecesario(rgOrd, outOrdValues, outOrdBgs);
    return;
  }

  const neededExam = (objective - achievedSoFar) / examW;

  if (neededExam - 1e-12 > maxGrade) {
    outOrdValues[1][0] = CFG.impossibleText;
    outOrdBgs[1][0] = CFG.whiteBg;
    writeNecesario(rgOrd, outOrdValues, outOrdBgs);
    return;
  }

  // Redondeo hacia arriba a 0.25 y respetar mínimo
  let clampedExam = clamp(neededExam, examMin, maxGrade);
  let roundedExam = roundUpToStep(clampedExam, step);

  // (Garantía) Si por redondeos/precisión no llegara, sube a base de step hasta llegar
  while (achievedSoFar + examW * roundedExam + 1e-12 < objective && roundedExam + step <= maxGrade) {
    roundedExam = Math.round((roundedExam + step) / step) * step;
  }

  outOrdValues[1][0] = roundedExam;
  outOrdBgs[1][0] = CFG.whiteBg;

  writeNecesario(rgOrd, outOrdValues, outOrdBgs);
}

/* ================= HELPERS ================= */

function writeNecesario(rg, values, bgs) {
  const n = values.length;
  const outColOffset = (CFG.colNecesario - 1);
  const outRange = rg.offset(0, outColOffset, n, 1);
  outRange.setValues(values);
  outRange.setBackgrounds(bgs);
}

function isBlank(v) {
  return v === "" || v === null || typeof v === "undefined";
}

function parseNumber(v) {
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    let s = v.trim().replace(/\s/g, "").replace(",", ".");
    s = s.replace("%", "");
    if (s === "") return null;
    const num = Number(s);
    if (!Number.isFinite(num)) return null;
    return num;
  }
  return null;
}

function clamp(x, lo, hi) {
  return Math.min(hi, Math.max(lo, x));
}

function roundUpToStep(x, step) {
  return Math.ceil(x / step) * step;
}/************ CONFIG ************/
const NEXT_EVENTS_ONOPEN_CFG = {
  sheetName: "database",

  // Source: start row and columns B..D (in sheetName)
  rowStart: 1497,
  colB: 2, // B
  colC: 3, // C
  colD: 4, // D

  // Output: hardcoded range J3:N7 (in MainDashboard)
  outputSheetName: "MainDashboard",
  outputStartRow: 3,
  outputStartCol: 10, // J
  outputNumRows: 5,
  outputNumCols: 5, // J..N

  // Only include events that start after now
  includeOngoing: false
};

/************ MAIN ************/
function writeNext5UpcomingEventsToJ3N7() {
  const cfg = NEXT_EVENTS_ONOPEN_CFG;

  const ss = SpreadsheetApp.getActive();

  // READ from database (cfg.sheetName)
  const shRead = ss.getSheetByName(cfg.sheetName);
  if (!shRead) throw new Error("Sheet not found (source): " + cfg.sheetName);

  // WRITE to MainDashboard
  const shWrite = ss.getSheetByName(cfg.outputSheetName);
  if (!shWrite) throw new Error("Sheet not found (output): " + cfg.outputSheetName);

  const now = new Date();

  // Clear output range J3:N7
  shWrite.getRange(cfg.outputStartRow, cfg.outputStartCol, cfg.outputNumRows, cfg.outputNumCols).clearContent();

  const lastRow = shRead.getLastRow();
  Logger.log("writeNext5UpcomingEventsToJ3N7: source lastRow=%s, rowStart=%s", lastRow, cfg.rowStart);

  if (lastRow < cfg.rowStart) {
    Logger.log("writeNext5UpcomingEventsToJ3N7: EXIT (source lastRow < rowStart). Nothing to read.");
    return;
  }

  // We will read BOTH potential blocks:
  // - Base block:   B/C/D  (cfg.colB, cfg.colB+1, cfg.colB+2)
  // - Alt block:    J/K/L  (B + 8, C + 8, D + 8)
  //
  // Note: although you mentioned "7 columns to the right", the explicit mapping you gave is B<->J,
  // which is +8 columns (B=2, J=10). We follow the explicit mapping: B and J, C and K, D and L.
  const ALT_OFFSET_COLS = 8;

  const numRows = lastRow - cfg.rowStart + 1;

  // Read one wide range from base B to alt L in one call for performance.
  // Width needed: base 3 cols + gap + alt 3 cols = ALT_OFFSET_COLS + 3 (because we start at cfg.colB).
  // Example if cfg.colB=2 (B) and ALT_OFFSET_COLS=8, width=11 => reads B..L.
  const width = ALT_OFFSET_COLS + 3;
  const grid = shRead.getRange(cfg.rowStart, cfg.colB, numRows, width).getValues();

  // Column indexes within `grid` (0-based inside the fetched range)
  const idxBaseB = 0;
  const idxBaseC = 1;
  const idxBaseD = 2;

  const idxAltJ = ALT_OFFSET_COLS + 0;
  const idxAltK = ALT_OFFSET_COLS + 1;
  const idxAltL = ALT_OFFSET_COLS + 2;

  const collected = [];

  // Helper: true if cell is empty (same semantics as your original code)
  function isEmpty_(v) {
    return !v || String(v).trim() === "";
  }

  // Helper: scan a single column independently until its first empty cell
  function scanColumnUntilEmpty_(colIdx, calendarType) {
    let count = 0;

    for (let i = 0; i < grid.length; i++) {
      const v = grid[i][colIdx];

      if (isEmpty_(v)) {
        // Stop condition for THIS column only: first empty cell
        Logger.log("Stopped %s at sheet row %s (first empty cell in that column).", calendarType, (cfg.rowStart + i));
        break;
      }

      const ev = parseLoggedCalendarCell_(String(v), calendarType);
      if (ev) {
        collected.push(ev);
        count++;
      }
    }

    Logger.log("Scanned %s columnIdx=%s: parsed=%s", calendarType, colIdx, count);
  }

  // Extract events equally from both blocks:
  // General:       B and J
  // Examenes:      C and K
  // Laboratorios:  D and L
  scanColumnUntilEmpty_(idxBaseB, "General");
  scanColumnUntilEmpty_(idxAltJ,  "General");

  scanColumnUntilEmpty_(idxBaseC, "Examenes");
  scanColumnUntilEmpty_(idxAltK,  "Examenes");

  scanColumnUntilEmpty_(idxBaseD, "Laboratorios");
  scanColumnUntilEmpty_(idxAltL,  "Laboratorios");

  Logger.log("Collected parsed events (combined): %s", collected.length);

  // Filter by time
  const filtered = collected.filter(ev => {
    if (!ev || !(ev.start instanceof Date) || isNaN(ev.start)) return false;

    if (cfg.includeOngoing) {
      if (ev.end instanceof Date && !isNaN(ev.end)) return ev.end.getTime() > now.getTime();
      return ev.start.getTime() > now.getTime();
    } else {
      return ev.start.getTime() > now.getTime();
    }
  });

  Logger.log("Filtered upcoming events: %s", filtered.length);

  // Sort by proximity (soonest start)
  filtered.sort((a, b) => a.start.getTime() - b.start.getTime());

  const top = filtered.slice(0, cfg.outputNumRows);

  // Build output matrix 5x5
  const out = [];
  for (let i = 0; i < cfg.outputNumRows; i++) {
    const ev = top[i];
    if (!ev) {
      out.push(["", "", "", "", ""]);
      continue;
    }

    const minutesUntil = Math.max(0, Math.round((ev.start.getTime() - now.getTime()) / 60000));

    out.push([
      ev.type,                          // Tipo
      ev.title,                         // Concepto
      formatDate_(ev.start),            // Fecha
      formatTime_(ev.start),            // Hora
      formatMinutesAsDHM_(minutesUntil) // T until
    ]);
  }

  // Write to J3:N7
  shWrite.getRange(cfg.outputStartRow, cfg.outputStartCol, cfg.outputNumRows, cfg.outputNumCols).setValues(out);
  Logger.log("Wrote %s rows to %s!J3:N7", out.length, cfg.outputSheetName);
}


/************ PARSER ************/
/**
 * Parses a cell block like:
 * [GOOGLE CALENDAR EVENT]
 * title: ...
 * id: ...
 * start: ...
 * end: ...
 */
function parseLoggedCalendarCell_(cellText, typeName) {
  const text = String(cellText || "").trim();
  if (!text) return null;

  const title = extractField_(text, "title");
  const id    = extractField_(text, "id");
  const startStr = extractField_(text, "start");
  const endStr   = extractField_(text, "end");

  const start = startStr ? new Date(startStr) : null;
  const end   = endStr ? new Date(endStr) : null;

  if (!(start instanceof Date) || isNaN(start)) return null;

  return {
    type: typeName,
    title: title || "",
    id: id || "",
    start: start,
    end: (end instanceof Date && !isNaN(end)) ? end : null
  };
}

function extractField_(text, fieldName) {
  const re = new RegExp("^\\s*" + escapeRegExp_(fieldName) + "\\s*:\\s*(.*)\\s*$", "mi");
  const m = text.match(re);
  return m ? (m[1] || "").trim() : "";
}

function escapeRegExp_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/************ FORMAT ************/
function formatDate_(d) {
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yy = d.getFullYear();
  return `${dd}/${mm}/${yy}`;
}

function formatTime_(d) {
  const hh = String(d.getHours()).padStart(2, "0");
  const mi = String(d.getMinutes()).padStart(2, "0");
  return `${hh}:${mi}`;
}

/**
 * Formats minutes as "Xd Yh Zm" (omits zero units except minutes if everything is zero).
 * Examples: "2d 3h 15m", "5h 20m", "12m"
 */
function formatMinutesAsDHM_(totalMinutes) {
  let mins = Math.max(0, totalMinutes);

  const days = Math.floor(mins / 1440);
  mins -= days * 1440;

  const hours = Math.floor(mins / 60);
  mins -= hours * 60;

  const parts = [];
  if (days > 0) parts.push(days + "d");
  if (hours > 0) parts.push(hours + "h");
  if (mins > 0 || parts.length === 0) parts.push(mins + "m");

  return parts.join(" ");
}
