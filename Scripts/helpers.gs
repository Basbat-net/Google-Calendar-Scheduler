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
        const idx = endRow - 1; // 1-based -> 0-based
        value = calA[idx]?.[0] ?? "";
      } catch (e) {
        value = "";
      }
    }

    output.push([value]);
  }

  shMain.getRange('G4:G22').setValues(output);

  // ─────────────────────────────────────────────
  // CONTADORES SEMANALES
  // ─────────────────────────────────────────────
  try {
    const countEstudio  = countTaggedSlots_WholeWeek_BtoO_("Estudio");
    const countProyecto = countTaggedSlots_WholeWeek_BtoO_("Proyecto");
    const countTareas   = countTaggedSlots_WholeWeek_BtoO_("Tareas");
    const countFromNowEstudio = countTaggedSlots_WholeWeek_BtoO_("Estudio", null);
    const countFromNowProyecto = countTaggedSlots_WholeWeek_BtoO_("Proyecto", null);
    const countFromNowTareas = countTaggedSlots_WholeWeek_BtoO_("Tareas", null);
    shMain.getRange('L11').setValue(countEstudio);
    shMain.getRange('L12').setValue(countProyecto);
    shMain.getRange('L13').setValue(countTareas);
    shMain.getRange('N11').setValue(countFromNowEstudio);
    shMain.getRange('N12').setValue(countFromNowProyecto);
    shMain.getRange('N13').setValue(countFromNowTareas);
  } catch (e) {
    shMain.getRange('L11:L13').setValues([[""], [""], [""]]);
  }
}





