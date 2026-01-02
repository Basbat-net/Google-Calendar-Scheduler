/**<--------------> FUNCIONES AUXILIARES <------------->**/

//REVISADO
function getNow_(extraOffset = 0) {return new Date((new Date()).getTime() + (NOW_OFFSET_HOURS + extraOffset) * 60 * 60 * 1000);}
  
//REVISADO
function startOfDay_(d) { return new Date(new Date(d).setHours(0, 0, 0, 0));}

//REVISADO
function addDays_(d, days) {return new Date(new Date(d).getTime() + days * 86400000);}

//REVISADO
function isSameDay_(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate(); }

//REVISADO
function computeIntervalForPriority_(examDate, priority) {
  const bracket = priority === 1 ? { min: 22, max: 365 } : { min: (5 - priority) * 7, max: (5 - priority) * 7 + 7};
  return { desde: addDays_(examDate, -bracket.max), hasta: addDays_(examDate, -bracket.min), priority }; }
//REVISADO
function getFillRatio_(task) {
  if(!task.deadline || task.isOverdue || (task.deadline - now) <= 0 || !task.etaMinutesRemaining || task.etaMinutesRemaining <= 0) return 0;
  return (task.etaMinutesRemaining / ((task.deadline - now) / 60000)); }

//REVISADO
function fetchCalendar_(name) {
  const existing = CalendarApp.getCalendarsByName(name);
  if (existing && existing.length > 0) return existing[0];
  return CalendarApp.createCalendar(name); }


//REVISADO
//Limpia hasta la fecha que le digas los eventos que le digas de un calendario (a partir de despues del actual evento)
function clearAutomaticEvents_(cal, startTime, endTime) {
  try {
    Logger.log("clearAutomaticEvents_: calendario '%s' (%s → %s)",
               cal.getName(), startTime, endTime);

    // Validaciones duras para que nunca llegue undefined al getEvents
    if (!(startTime instanceof Date) || isNaN(startTime.getTime())) {
      Logger.log("clearAutomaticEvents_: startTime inválido: %s", startTime);
      return;
    }
    if (!(endTime instanceof Date) || isNaN(endTime.getTime())) {
      Logger.log("clearAutomaticEvents_: endTime inválido: %s", endTime);
      return;
    }

    // Normalizar por si vienen invertidas
    if (endTime.getTime() <= startTime.getTime()) {
      Logger.log("clearAutomaticEvents_: rango inválido (endTime <= startTime). No se borra nada.");
      return;
    }

    const events = cal.getEvents(startTime, endTime);
    Logger.log("clearAutomaticEvents_: encontrados %s eventos para borrar.", events.length);

    // ✅ BORRADO TOTAL de todos los eventos en el rango
    let deleted = 0;
    for (let i = 0; i < events.length; i++) {
      try {
        events[i].deleteEvent();
        deleted++;
      } catch (evErr) {
        Logger.log("clearAutomaticEvents_: error borrando evento #%s (%s): %s",
                   i + 1, safeEventDebugName_(events[i]), evErr);
      }
    }

    Logger.log("clearAutomaticEvents_: borrados %s/%s eventos.", deleted, events.length);

  } catch (e) {
    Logger.log("getEvents falló en calendario '%s' (%s → %s): %s",
               (cal && cal.getName) ? cal.getName() : cal, startTime, endTime, e);
  }
}

// Helper para logs seguros (evita que un fallo al leer el título rompa el borrado)
function safeEventDebugName_(ev) {
  try {
    const title = ev.getTitle ? ev.getTitle() : "";
    const s = ev.getStartTime ? ev.getStartTime() : "";
    const e = ev.getEndTime ? ev.getEndTime() : "";
    return `title="${title}" start=${s} end=${e}`;
  } catch (_) {
    return "(evento: sin detalles)";
  }
}

//REVISADO
function dateKeyFromDate_(d){return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);}

//REVISADO
function dateFromKey_(key) {
  const [y, m, d] = key.split('-').map(Number);
  return new Date(y, m - 1, d);
}

//REVISADO
function roundUpToNextHalfHour_(d) {
  const res = new Date(d);
  res.setSeconds(0, 0);
  const m = res.getMinutes();    
  res.setMinutes(m < 30 ? 30 : 0);
  if (res.getMinutes() >= 30) res.setHours(res.getHours() + 1);
  return res;
}

//REVISADO
function createMergedEventsForSegments_(cal, title, segments) {
  if (!cal || !segments || segments.length === 0) return;

// --- LOGGING TO SHEETS (database) BEFORE calendar writes ---
// Column mapping by calendar name (as requested)
const DB_SHEET_NAME = "database";
const START_ROW = 1497;

const COL_BY_CALNAME = {
  "Tareas": "N",
  "Estudio": "O",
  "Proyecto": "P"
};

const db = SpreadsheetApp.getActive().getSheetByName(DB_SHEET_NAME);
if (db) {
  const tz = Session.getScriptTimeZone();
  const calName = String(cal.getName ? cal.getName() : "");
  const colLetter = COL_BY_CALNAME[calName];

  function colLetterToNumber_(letter) {
    let n = 0;
    for (let i = 0; i < letter.length; i++) n = n * 26 + (letter.charCodeAt(i) - 64);
    return n;
  }

  // Finds last non-empty row (>= startRow) in a single column; returns startRow-1 if empty
  function findLastNonEmptyRowInCol_(sheet, colLetter, startRow) {
    const col = colLetterToNumber_(colLetter);
    const maxRows = sheet.getMaxRows();
    const n = Math.max(0, maxRows - startRow + 1);
    if (n <= 0) return startRow - 1;

    const vals = sheet.getRange(startRow, col, n, 1).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      const v = vals[i][0];
      if (!(v === "" || v === null)) return startRow + i;
    }
    return startRow - 1;
  }

  // --- NEW: clear the target column (from START_ROW to last non-blank) every execution ---
  if (colLetter) {
    const lastUsedRow = findLastNonEmptyRowInCol_(db, colLetter, START_ROW);
    if (lastUsedRow >= START_ROW) {
      const col = colLetterToNumber_(colLetter);
      db.getRange(START_ROW, col, lastUsedRow - START_ROW + 1, 1).clearContent();
    }
  }

  if (colLetter) {
    // Build logs (one row per merged event that will be created)
    const filteredAllDays = segments
      .filter(seg => seg && seg.start instanceof Date && seg.end instanceof Date)
      .sort((a, b) => a.start.getTime() - b.start.getTime());

    if (filteredAllDays.length > 0) {
      // Merge segments with the same logic you use for calendar creation
      const mergedRanges = [];
      let cur = {
        start: new Date(filteredAllDays[0].start.getTime()),
        end:   new Date(filteredAllDays[0].end.getTime())
      };

      for (let i = 1; i < filteredAllDays.length; i++) {
        const seg = filteredAllDays[i];
        if (seg.start.getTime() <= cur.end.getTime() + 10000) {
          if (seg.end.getTime() > cur.end.getTime()) {
            cur.end = new Date(seg.end.getTime());
          }
        } else {
          mergedRanges.push(cur);
          cur = { start: new Date(seg.start.getTime()), end: new Date(seg.end.getTime()) };
        }
      }
      mergedRanges.push(cur);

      const rowsToWrite = mergedRanges.map(r => {
        const startStr = Utilities.formatDate(r.start, tz, "EEE MMM dd yyyy HH:mm:ss 'GMT'Z");
        const endStr   = Utilities.formatDate(r.end,   tz, "EEE MMM dd yyyy HH:mm:ss 'GMT'Z");

        const logText =
          "[GOOGLE CALENDAR EVENT]\n" +
          "title: " + title + "\n" +
          "id: \n" + // ID no existe aún (se crea tras createEvent)
          "start: " + startStr + "\n" +
          "end: " + endStr + "\n" +
          "allDay: false\n" +
          "location: \n" +
          "created: \n" +
          "lastUpdated: \n" +
          "description:\n";

        return [logText];
      });

      // After clearing, always write starting at START_ROW
      const col = colLetterToNumber_(colLetter);

      const neededLastRow = START_ROW + rowsToWrite.length - 1;
      const currentMax = db.getMaxRows();
      if (neededLastRow > currentMax) {
        db.insertRowsAfter(currentMax, neededLastRow - currentMax);
      }

      db.getRange(START_ROW, col, rowsToWrite.length, 1).setValues(rowsToWrite);
    }
  }
}
// --- END LOGGING ---

  // --- NORMAL CALENDAR CREATION FOR ALL DAYS ---
  const filtered = segments
    .filter(seg => seg && seg.start instanceof Date && seg.end instanceof Date)
    .sort((a, b) => a.start.getTime() - b.start.getTime());

  if (filtered.length === 0) return;

  let current = {
    start: new Date(filtered[0].start.getTime()),
    end:   new Date(filtered[0].end.getTime())
  };

  for (let i = 1; i < filtered.length; i++) {
    const seg = filtered[i];

    if (seg.start.getTime() <= current.end.getTime() + 10000) {
      if (seg.end.getTime() > current.end.getTime()) {
        current.end = new Date(seg.end.getTime());
      }
    } else {
      cal.createEvent(title, current.start, current.end);
      current = {
        start: new Date(seg.start.getTime()),
        end:   new Date(seg.end.getTime())
      };
    }
  }

  cal.createEvent(title, current.start, current.end);

  }




function renderMainDashboard_G4_values() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shMain = ss.getSheetByName('MainDashboard');
  const shCal  = ss.getSheetByName('Calendario');

  if (!shMain) throw new Error("No existe 'MainDashboard'");
  if (!shCal)  throw new Error("No existe 'Calendario'");

  // Rangos equivalentes a la fórmula
  const startRow = 53;
  const numRows  = 19; // 53..71

  // ─────────────────────────────────────────────
  // ✅ Sustituto de MATCH (C53:C71) y VLOOKUP (hora en A)
  // ─────────────────────────────────────────────
  const names = shMain
    .getRange(`B${startRow}:B${startRow + numRows - 1}`)
    .getValues()
    .map(r => r[0]);

  const baseColIndex = Number(shMain.getRange('G53').getValue()); // $G$53 en la fórmula

  // Columna A del Calendario (para devolver la hora)
  const calA = shCal.getRange('A:A').getValues();

  let startRows = new Array(numRows).fill("");
  let startTimes = new Array(numRows).fill("");

  if (baseColIndex && baseColIndex > 0) {
    const lastRow = shCal.getLastRow();
    const height = Math.max(0, lastRow - 1); // desde fila 2

    const col1Vals = height ? shCal.getRange(2, baseColIndex, height, 1).getValues() : [];
    const col2Vals = height ? shCal.getRange(2, baseColIndex + 1, height, 1).getValues() : [];

    const usedRowSet = new Set(); // filas (1-based) ya asignadas

    function findNextStartRow_(needleRaw, columnValues, colIndex) {
      const needle = String(needleRaw ?? "").trim().toLowerCase();
      if (!needle) return "";

      for (let i = 0; i < columnValues.length; i++) {
        const v = columnValues[i][0];
        if (v === "" || v === null || v === undefined) continue;

        const text = String(v).trim().toLowerCase();
        if (!text.startsWith(needle)) continue;

        const realRow = i + 2; // array empieza en fila 2

        if (usedRowSet.has(realRow)) continue;

        // Asegurar que es el inicio del bloque (top-left) si está mergeado
        const cell = shCal.getRange(realRow, colIndex, 1, 1);
        if (cell.isPartOfMerge()) {
          const merges = cell.getMergedRanges();
          if (merges && merges.length) {
            let ok = false;
            for (let k = 0; k < merges.length; k++) {
              const m = merges[k];
              const r0 = m.getRow();
              const c0 = m.getColumn();
              const r1 = r0 + m.getNumRows() - 1;
              const c1 = c0 + m.getNumColumns() - 1;

              if (realRow >= r0 && realRow <= r1 && colIndex >= c0 && colIndex <= c1) {
                if (realRow === r0 && colIndex === c0) ok = true;
                break;
              }
            }
            if (!ok) continue;
          }
        }

        usedRowSet.add(realRow);
        return realRow;
      }

      return "";
    }

    startRows = names.map(name => {
      if (name === "" || name === null || name === undefined) return "";
      const r1 = findNextStartRow_(name, col1Vals, baseColIndex);
      if (r1 !== "") return r1;
      const r2 = findNextStartRow_(name, col2Vals, baseColIndex + 1);
      if (r2 !== "") return r2;
      return "";
    });

    startTimes = startRows.map(r => {
      if (r === "" || r === null || r === undefined) return "";
      const idx = Number(r) - 1; // 1-based -> 0-based
      return calA[idx]?.[0] ?? "";
    });
  }

  // Si quieres seguir sustituyendo la fórmula original de C53:C71, esto se queda:
  shMain
    .getRange(`C${startRow}:C${startRow + numRows - 1}`)
    .setValues(startRows.map(r => [r]));

  // ✅ AQUÍ ESTÁ EL FIX: escribir en el dashboard desde F4 (no F53)
  shMain
    .getRange('F4:F22')
    .setValues(startTimes.map(v => [v]));

  // ─────────────────────────────────────────────
  // Tu lógica original para G4:G22 (fin del merge)
  // ─────────────────────────────────────────────
  const N = shMain.getRange(`N${startRow}:N${startRow + numRows - 1}`).getValues();
  const D = shMain.getRange(`D${startRow}:D${startRow + numRows - 1}`).getValues();
  const C = shMain.getRange(`C${startRow}:C${startRow + numRows - 1}`).getValues();

  const output = [];

  for (let i = 0; i < numRows; i++) {
    const n   = N[i][0];
    const col = D[i][0];
    const r   = C[i][0];

    let value = "";

    if (n !== "") {
      value = "";
    } else if (col === "") {
      value = "";
    } else {
      try {
        const endRow = getMergeEnd("Calendario", col, r);
        const idx = endRow; // 1-based -> 0-based
        value = calA[idx]?.[0] ?? "";
      } catch (e) {
        value = "";
      }
    }

    output.push([value]);
  }

  shMain.getRange('G4:G22').setValues(output);

}


/**
 * Para cada concepto en K11:K18:
 *  - Suma semanal (desde el último lunes inclusive) de Estudio en L11:L18
 *  - Suma total histórica (sin filtro de fecha) de Estudio en N11:N18
 *
 * Extra:
 *  - Suma semanal TOTAL de Estudio (todos los conceptos) -> J14
 *  - Suma total histórica de Estudio (todos los conceptos) -> J18
 *
 * NUEVO:
 *  - Suma semanal TOTAL de Proyecto -> L19
 *  - Suma total histórica de Proyecto -> N19
 *  - Suma semanal TOTAL de Tareas -> L20
 *  - Suma total histórica de Tareas -> N20
 *
 * ANTES DE ESCRIBIR RESULTADOS:
 *  - Lee 'MainDashboard'!D4:H49
 *  - Recorta por la primera fila vacía en columna D
 *  - Cuenta intervalos de 30 min para:
 *      * Proyecto / Tareas (por tipo)
 *      * Estudio TOTAL (por tipo)
 *      * Estudio POR ASIGNATURA: tipo=Estudio y concepto (col D) ∈ K11:K18
 *
 * AHORA:
 *  - SUMA (MainDashboard + Database) y vuelca la SUMA al sheet (archivo)
 *
 * DATABASE ('database'!A1001:F1491):
 *  - A: Fecha
 *  - B: Concepto
 *  - C: Tipo  (se filtra por "Estudio" / "Proyecto" / "Tareas")
 *  - F: Duración (horas)
 */
function updateEstudioWeekAndTotalSums() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  const db = ss.getSheetByName("database");
  if (!db) throw new Error("No existe la hoja 'database'.");

  /***********************
   * 1) INPUT CONCEPTOS (para DB y para dashboard por-asignatura)
   ***********************/
  const START_ROW = 11;
  const END_ROW   = 18;
  const numRows   = END_ROW - START_ROW + 1;

  // K = 11
  const conceptos = sh
    .getRange(START_ROW, 11, numRows, 1) // K11:K18
    .getValues()
    .flat()
    .map(v => (v || "").toString().trim());

  const conceptosSet = new Set(conceptos.filter(Boolean));

  /***********************
   * 0) PRE-CHECK DASHBOARD (contadores en HORAS)
   ***********************/
  const shMain = ss.getSheetByName("MainDashboard");
  if (!shMain) throw new Error("No existe la hoja 'MainDashboard'.");

  const DASH_START_ROW = 4;
  const DASH_END_ROW   = 49;
  const dashNumRows = DASH_END_ROW - DASH_START_ROW + 1;

  // D:E:F:G:H -> Concepto, Tipo, Inicio, Final, Done?
  const dashVals = shMain.getRange(DASH_START_ROW, 4, dashNumRows, 5).getValues();

  function norm_(v) {
    return (v == null ? "" : String(v))
      .replace(/\u00A0/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  // primera fila vacía en D
  let firstEmptyIdx = -1;
  for (let i = 0; i < dashVals.length; i++) {
    const conceptoCell = norm_(dashVals[i][0]); // col D
    if (!conceptoCell) {
      firstEmptyIdx = i;
      break;
    }
  }
  const effectiveLen = (firstEmptyIdx === -1) ? dashVals.length : firstEmptyIdx;

  const HALF_HOUR_MS = 30 * 60 * 1000;

  function intervals30mBetween_(startVal, endVal) {
    if (!(startVal instanceof Date) || !(endVal instanceof Date)) return 0;

    let diffMs = endVal.getTime() - startVal.getTime();
    if (diffMs < 0) diffMs += 24 * 60 * 60 * 1000;

    const intervals = Math.round(diffMs / HALF_HOUR_MS);
    if (!isFinite(intervals) || intervals < 0) return 0;
    return intervals;
  }

  // Contadores del dashboard en INTERVALOS
  let dashEstudioIntervals  = 0;
  let dashProyectoIntervals = 0;
  let dashTareasIntervals   = 0;

  // Estudio por asignatura (intervalos)
  const dashEstudioBySubjectIntervals = {};
  for (const c of conceptos) {
    const key = (c || "").toString().trim();
    if (key) dashEstudioBySubjectIntervals[key] = 0;
  }

  Logger.log("---- DASHBOARD DEBUG (MainDashboard D4:H49) effectiveLen=%s ----", effectiveLen);

  for (let i = 0; i < effectiveLen; i++) {
    const row = dashVals[i];

    const conceptoRaw = row[0]; // D
    const tipoRaw     = row[1]; // E
    const ini         = row[2]; // F
    const fin         = row[3]; // G
    const doneRaw     = row[4]; // H

    const concepto = norm_(conceptoRaw);
    const tipo     = norm_(tipoRaw);
    const done = (doneRaw === true || doneRaw === "TRUE" || doneRaw === 1);

    let intervals = 0;

    if (done && (tipo === "Estudio" || tipo === "Proyecto" || tipo === "Tareas")) {
      intervals = intervals30mBetween_(ini, fin);

      if (tipo === "Estudio") {
        dashEstudioIntervals += intervals;

        if (conceptosSet.has(concepto)) {
          if (dashEstudioBySubjectIntervals[concepto] == null) dashEstudioBySubjectIntervals[concepto] = 0;
          dashEstudioBySubjectIntervals[concepto] += intervals;
        }
      } else if (tipo === "Proyecto") {
        dashProyectoIntervals += intervals;
      } else if (tipo === "Tareas") {
        dashTareasIntervals += intervals;
      }
    }

    Logger.log(
      "Row %s: concepto='%s' | tipoRaw='%s' (norm='%s') | doneRaw=%s | ini=%s | fin=%s | intervals=%s",
      DASH_START_ROW + i,
      concepto,
      tipoRaw,
      tipo,
      doneRaw,
      ini,
      fin,
      intervals
    );
  }

  // Convertimos a HORAS (0.5h por intervalo)
  const dashEstudioHours  = dashEstudioIntervals * 0.5;
  const dashProyectoHours = dashProyectoIntervals * 0.5;
  const dashTareasHours   = dashTareasIntervals * 0.5;

  const dashEstudioBySubjectHours = {};
  for (const subj of conceptos) {
    const key = (subj || "").toString().trim();
    if (!key) continue;
    const ints = Number(dashEstudioBySubjectIntervals[key]) || 0;
    dashEstudioBySubjectHours[key] = ints * 0.5;
  }

  Logger.log(
    "MainDashboard totals: Estudio=%s intervalos (%s h), Proyecto=%s intervalos (%s h), Tareas=%s intervalos (%s h)",
    dashEstudioIntervals,  dashEstudioHours.toFixed(2),
    dashProyectoIntervals, dashProyectoHours.toFixed(2),
    dashTareasIntervals,   dashTareasHours.toFixed(2)
  );

  Logger.log("MainDashboard Estudio por asignatura (solo concepto ∈ K11:K18):");
  for (const subj of conceptos) {
    const key = (subj || "").toString().trim();
    if (!key) continue;
    const h = Number(dashEstudioBySubjectHours[key]) || 0;
    Logger.log("  - %s: %s h", key, h.toFixed(2));
  }

  /***********************
   * 2) DB: Ventana temporal desde último lunes
   ***********************/
  const now = new Date();

  const minDate = new Date(now);
  const dow = minDate.getDay();
  const daysSinceMonday = (dow + 6) % 7; // Lun->0, Mar->1, ..., Dom->6
  minDate.setDate(minDate.getDate() - daysSinceMonday);
  minDate.setHours(0, 0, 0, 0);

  const dbValues = db.getRange("A1001:F1491").getValues();

  // totales globales DB (horas)
  let totalEstudioWeek_allConcepts_DB = 0; // -> J14 (si lo usas)
  let totalEstudioAll_allConcepts_DB  = 0; // -> J18 (si lo usas)

  let totalProyectoWeek_all_DB = 0; // -> L19
  let totalProyectoAll_all_DB  = 0; // -> N19
  let totalTareasWeek_all_DB   = 0; // -> L20
  let totalTareasAll_all_DB    = 0; // -> N20

  // --- calcular por concepto (Estudio) en DB ---
  const weeklyOut = [];
  const totalOut  = [];

  // aquí guardamos también los sumatorios DB por asignatura para logging final
  const estudioWeekBySubject_DB = {};
  const estudioAllBySubject_DB  = {};
  for (const c of conceptos) {
    const key = (c || "").toString().trim();
    if (key) {
      estudioWeekBySubject_DB[key] = 0;
      estudioAllBySubject_DB[key] = 0;
    }
  }

  for (const concepto of conceptos) {
    if (!concepto) {
      weeklyOut.push([""]);
      totalOut.push([""]);
      continue;
    }

    let sumWeek_DB = 0;
    let sumAll_DB  = 0;

    for (const row of dbValues) {
      const fecha      = row[0]; // A
      const conceptoDB = (row[1] || "").toString().trim(); // B
      const tipo       = row[2]; // C
      const duracion   = Number(row[5]) || 0; // F (horas)

      if (tipo !== "Estudio") continue;
      if (conceptoDB !== concepto) continue;

      sumAll_DB += duracion;

      if (fecha instanceof Date && fecha >= minDate && fecha <= now) {
        sumWeek_DB += duracion;
      }
    }

    // Guardamos para log
    estudioWeekBySubject_DB[concepto] = sumWeek_DB;
    estudioAllBySubject_DB[concepto]  = sumAll_DB;

    // Totales globales de Estudio (DB)
    totalEstudioAll_allConcepts_DB  += sumAll_DB;
    totalEstudioWeek_allConcepts_DB += sumWeek_DB;

    // --- SUMA: DB + MainDashboard (por asignatura) ---
    // Nota: los bloques del dashboard no tienen fecha real (salen en 1899), así que por diseño
    // los sumamos tanto a semana como a total (son “no logueados” aún).
    const dashH = Number(dashEstudioBySubjectHours[concepto]) || 0;

    const sumWeek_TOTAL = sumWeek_DB + dashH;
    const sumAll_TOTAL  = sumAll_DB + dashH;

    weeklyOut.push([sumWeek_TOTAL]);
    totalOut.push([sumAll_TOTAL]);
  }

  /***********************
   * 3) SUMA: DB + MainDashboard (Proyecto / Tareas totales)
   ***********************/
  for (const row of dbValues) {
    const fecha    = row[0]; // A
    const tipo     = (row[2] || "").toString().trim(); // C
    const duracion = Number(row[5]) || 0; // F (horas)

    const inThisWeek = (fecha instanceof Date && fecha >= minDate && fecha <= now);

    if (tipo === "Proyecto") {
      totalProyectoAll_all_DB += duracion;
      if (inThisWeek) totalProyectoWeek_all_DB += duracion;
    } else if (tipo === "Tareas") {
      totalTareasAll_all_DB += duracion;
      if (inThisWeek) totalTareasWeek_all_DB += duracion;
    }
  }

  // SUMA: DB + Dashboard (en horas)
  // Igual que con Estudio: se suman a semana y a total porque son “no logueados aún”.
  const totalProyectoWeek_TOTAL = totalProyectoWeek_all_DB + dashProyectoHours;
  const totalProyectoAll_TOTAL  = totalProyectoAll_all_DB  + dashProyectoHours;

  const totalTareasWeek_TOTAL = totalTareasWeek_all_DB + dashTareasHours;
  const totalTareasAll_TOTAL  = totalTareasAll_all_DB  + dashTareasHours;

  const totalEstudioWeek_allConcepts_TOTAL = totalEstudioWeek_allConcepts_DB + dashEstudioHours;
  const totalEstudioAll_allConcepts_TOTAL  = totalEstudioAll_allConcepts_DB  + dashEstudioHours;

  /***********************
   * 4) OUTPUT (archivo / sheet)
   ***********************/
  // Por asignatura
  sh.getRange(START_ROW, 12, numRows, 1).setValues(weeklyOut); // L11:L18 (DB + Dashboard)
  sh.getRange(START_ROW, 14, numRows, 1).setValues(totalOut);  // N11:N18 (DB + Dashboard)

  // Totales Proyecto/Tareas (DB + Dashboard)
  sh.getRange("L19").setValue(totalProyectoWeek_TOTAL);
  sh.getRange("N19").setValue(totalProyectoAll_TOTAL);
  sh.getRange("L20").setValue(totalTareasWeek_TOTAL);
  sh.getRange("N20").setValue(totalTareasAll_TOTAL);

  // Si quieres también los totales de Estudio (DB + Dashboard), descomenta:
  // sh.getRange("J14").setValue(totalEstudioWeek_allConcepts_TOTAL);
  // sh.getRange("J18").setValue(totalEstudioAll_allConcepts_TOTAL);

  /***********************
   * 5) LOG FINAL (para verificar sumas)
   ***********************/
  Logger.log("---- SUMA FINAL (DB + MainDashboard) ----");
  Logger.log(
    "Estudio TOTAL: semana=%s h, total=%s h",
    totalEstudioWeek_allConcepts_TOTAL.toFixed(2),
    totalEstudioAll_allConcepts_TOTAL.toFixed(2)
  );
  Logger.log(
    "Proyecto TOTAL: semana=%s h, total=%s h",
    totalProyectoWeek_TOTAL.toFixed(2),
    totalProyectoAll_TOTAL.toFixed(2)
  );
  Logger.log(
    "Tareas TOTAL: semana=%s h, total=%s h",
    totalTareasWeek_TOTAL.toFixed(2),
    totalTareasAll_TOTAL.toFixed(2)
  );

  Logger.log("Estudio por asignatura (DB + MainDashboard):");
  for (const subj of conceptos) {
    const key = (subj || "").toString().trim();
    if (!key) continue;

    const dbW = Number(estudioWeekBySubject_DB[key]) || 0;
    const dbA = Number(estudioAllBySubject_DB[key]) || 0;
    const dh  = Number(dashEstudioBySubjectHours[key]) || 0;

    Logger.log(
      "  - %s: semana=%s h (DB=%s + Dash=%s), total=%s h (DB=%s + Dash=%s)",
      key,
      (dbW + dh).toFixed(2), dbW.toFixed(2), dh.toFixed(2),
      (dbA + dh).toFixed(2), dbA.toFixed(2), dh.toFixed(2)
    );
  }
}

