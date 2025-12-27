/************ GLOBAL STATE ************/

// Contendrá los logs completos leídos desde database (AHORA desde A..D)
// Estructura:
// {
//   Clases:        [string, ...],
//   General:       [string, ...],
//   Examenes:      [string, ...],
//   Laboratorios:  [string, ...]
// }
let LOGGED_CALENDAR_EVENTS = {};

// NUEVO: cuántos logs se han movido en el último CUT desde I/J/K/L -> A/B/C/D
let LAST_CALENDAR_WRITE_MOVED = 0;

/**
 * 1) Primero hace CUT: mueve (append) todo lo que haya en I..L (WRITE) hacia A..D (CHECK),
 *    limpiando la fuente y compactando.
 * 2) Luego lee los logs COMPLETOS desde A..D (CHECK) y los vuelca en LOGGED_CALENDAR_EVENTS.
 */

function readLoggedCalendarEvents_() {
  const CALENDAR_WRITE_COL = {
    "Clases":        "I",
    "General":       "J",
    "Examenes":      "K",
    "Laboratorios":  "L"
  };

  const CALENDAR_CHECK_COL = {
    "Clases":        "A",
    "General":       "B",
    "Examenes":      "C",
    "Laboratorios":  "D"
  };

  const START_ROW = 1497;
  const END_ROW   = 2000;

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("database");
  if (!sheet) {
    throw new Error("No existe la hoja 'database'");
  }

  // NUEVO: reset del contador por ejecución
  let movedTotal = 0;

  // ============================================================
  // 1) CUT: mover desde WRITE_COL (I..L) hacia CHECK_COL (A..D)
  // ============================================================
  Object.entries(CALENDAR_WRITE_COL).forEach(([calendarName, srcColLetter]) => {
    const srcRangeA1 = `${srcColLetter}${START_ROW}:${srcColLetter}${END_ROW}`;
    const srcValues = sheet.getRange(srcRangeA1).getDisplayValues().flat();

    const cleanedSrc = srcValues
      .map(v => (v == null ? "" : String(v).trim()))
      .filter(v => v !== "");

    if (cleanedSrc.length === 0) return;

    const destColLetter = CALENDAR_CHECK_COL[calendarName];
    if (!destColLetter) return;

    const destRangeA1 = `${destColLetter}${START_ROW}:${destColLetter}${END_ROW}`;
    const destVals = sheet.getRange(destRangeA1).getDisplayValues().flat();

    let lastNonEmptyRel = -1;
    for (let i = destVals.length - 1; i >= 0; i--) {
      const v = destVals[i];
      if (v != null && String(v).trim() !== "") {
        lastNonEmptyRel = i;
        break;
      }
    }

    let writeRow = START_ROW + (lastNonEmptyRel + 1);
    const capacity = (END_ROW - writeRow + 1);
    if (capacity <= 0) {
      Logger.log("readLoggedCalendarEvents_: WARNING -> No hay espacio en " + destRangeA1 + " para mover (" + calendarName + ")");
      return;
    }

    const toMove = cleanedSrc.slice(0, capacity);
    if (toMove.length === 0) return;

    // NUEVO: acumula lo movido
    movedTotal += toMove.length;

    const out = toMove.map(x => [x]);
    sheet.getRange(writeRow, columnLetterToIndex_(destColLetter), out.length, 1).setValues(out);

    Logger.log("readLoggedCalendarEvents_: CUT moved " + toMove.length + " logs de " + srcColLetter + " -> " + destColLetter +
               " (" + destColLetter + writeRow + ":" + destColLetter + (writeRow + toMove.length - 1) + ")");

    sheet.getRange(START_ROW, columnLetterToIndex_(srcColLetter), toMove.length, 1).clearContent();
    compactColumnUpInRange_(sheet, srcColLetter, START_ROW, END_ROW);
  });

  // NUEVO: expone el resultado del CUT
  LAST_CALENDAR_WRITE_MOVED = movedTotal;


  // ============================================================
  // 2) READ: leer AHORA desde CHECK_COL (A..D) y set global
  // ============================================================
  const result = {};
  let anyFound = false;

  Object.entries(CALENDAR_CHECK_COL).forEach(([calendarName, colLetter]) => {
    const rangeA1 = `${colLetter}${START_ROW}:${colLetter}${END_ROW}`;

    const values = sheet.getRange(rangeA1).getDisplayValues();

    const cleaned = values
      .flat()
      .map(v => (v == null ? "" : String(v).trim()))
      .filter(v => v !== "");

    result[calendarName] = cleaned;
    if (cleaned.length > 0) anyFound = true;

    Logger.log("readLoggedCalendarEvents_: READ " + calendarName + " -> " + cleaned.length + " logs (rango " + rangeA1 + ")");
    if (cleaned.length > 0) {
      Logger.log("readLoggedCalendarEvents_: ejemplo " + calendarName + " -> " + cleaned[0].slice(0, 200));
    }
  });

  LOGGED_CALENDAR_EVENTS = result;

  if (!anyFound) {
    Logger.log("readLoggedCalendarEvents_: WARNING -> No se encontró ningún log no vacío en A/B/C/D filas 1497..2000.");
    Logger.log("readLoggedCalendarEvents_: Verifica que realmente están en database!A1497:A2000, B..., C..., D...");
  }
}


/**
 * Compacta una columna dentro de un rango [startRow..endRow]:
 * - elimina huecos (celdas vacías intermedias)
 * - mantiene el orden relativo de los no vacíos
 * - NO escribe fuera del intervalo
 */
function compactColumnUpInRange_(sheet, colLetter, startRow, endRow) {
  const colIndex = columnLetterToIndex_(colLetter);
  const numRows = endRow - startRow + 1;

  const range = sheet.getRange(startRow, colIndex, numRows, 1);
  const values = range.getDisplayValues().flat();

  const kept = values
    .map(v => (v == null ? "" : String(v).trim()))
    .filter(v => v !== "");

  const out = [];
  for (let i = 0; i < numRows; i++) {
    out.push([ i < kept.length ? kept[i] : "" ]);
  }

  range.setValues(out);
}


/**
 * Convierte letra(s) de columna (A, B, Z, AA, AB...) a índice 1-based.
 */
function columnLetterToIndex_(letters) {
  let n = 0;
  const s = String(letters).toUpperCase().trim();
  for (let i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n;
}

// REVISADO
function getFixedEventsFromLoggedDatabase_(start, end) {
  // Ahora: readLoggedCalendarEvents_ hace CUT primero y luego LEE desde A..D
  try {
    readLoggedCalendarEvents_();
  } catch (e) {
    Logger.log("getFixedEventsFromLoggedDatabase_: no se pudo cargar LOGGED_CALENDAR_EVENTS: " + e);
    return [];
  }

  const busy = [];

  function extractField_(logStr, field) {
    if (logStr == null) return null;
    const s = String(logStr);

    // 1) Multilínea: start: ... / end: ...
    let m = s.match(new RegExp("^\\s*" + field + "\\s*:\\s*(.+)\\s*$", "im"));
    if (m && m[1]) return m[1].trim();

    // 2) JSON: "start": "..."
    m = s.match(new RegExp('"' + field + '"\\s*:\\s*"([^"]+)"', "i"));
    if (m && m[1]) return m[1].trim();

    // 3) start=...
    m = s.match(new RegExp("(^|\\s|[,{])" + field + "\\s*=\\s*([^\\n\\r,}\\]]+)", "i"));
    if (m && m[2]) return m[2].trim();

    // 4) start: ... inline
    m = s.match(new RegExp("(^|\\s|[,{])" + field + "\\s*:\\s*([^\\n\\r,}\\]]+)", "i"));
    if (m && m[2]) return m[2].trim();

    return null;
  }

  function parseStartEndFromLog_(logStr) {
    const startRaw = extractField_(logStr, "start");
    const endRaw   = extractField_(logStr, "end");
    if (!startRaw || !endRaw) return null;

    const st = new Date(Date.parse(startRaw));
    const en = new Date(Date.parse(endRaw));

    if (isNaN(st.getTime()) || isNaN(en.getTime())) return null;
    return { start: st, end: en };
  }

  let parsedOk = 0;
  let parsedFail = 0;

  Object.entries(LOGGED_CALENDAR_EVENTS).forEach(([calName, logs]) => {
    (logs || []).forEach(logStr => {
      const parsed = parseStartEndFromLog_(logStr);
      if (!parsed) {
        parsedFail++;
        return;
      }
      parsedOk++;

      if (parsed.start < end && parsed.end > start) {
        busy.push({
          cal: calName,
          start: parsed.start,
          end: parsed.end,
          raw: String(logStr)
        });
      }
    });
  });

  Logger.log("getFixedEventsFromLoggedDatabase_: parsedOk=" + parsedOk + " parsedFail=" + parsedFail + " busyEnRango=" + busy.length);

  return busy;
}

// REVISADO
function buildFreeSlots_(start, end){
  const rangeStart = roundUpToNextHalfHour_(start);
  const rangeEnd   = roundUpToNextHalfHour_(end);

  // Sustituye getFixedEvents_(): ahora los "busy" salen de los logs en database (I..L, filas 1497:2000)
  const fixedBusyEvents = getFixedEventsFromLoggedDatabase_(rangeStart, rangeEnd); // aqui sacamos lo que no podemos tocar

  // Ahora iremos metiendo directamente los huecos ya fusionados
  const mergedFreeSlots = [];

  // Este bucle es el que mete los intervalos de media hora en el array
  for (let d = new Date(rangeStart); d < rangeEnd; d = new Date(d.getTime() + (24 * 60 * 60 * 1000))) {
    // itera cada dia
    const day = new Date(d.getFullYear(), d.getMonth(), d.getDate());

    // IMPORTANTE: Los intervalos son continuos, no se parten
    const dow = day.getDay(); // 0 = Sunday, 6 = Saturday
    const blocks = (dow === 0 || dow === 6)? INTERVALOS_USABLES_FINDES: INTERVALOS_USABLES_ENTRE_SEMANA;

    // Itera todo el dia en intervalos de media hora
    blocks.forEach(block => {
      let blockStart = new Date(day.getFullYear(), day.getMonth(), day.getDate(), block.startHour, block.startMinute);
      let blockEnd = new Date(day.getFullYear(), day.getMonth(), day.getDate(), block.endHour, block.endMinute);

      // para ignorar cosas fuera de los horarios de curro
      if (blockEnd <= rangeStart || blockStart >= rangeEnd) return;

      // filtra para que las cosas te caigan en el intervalo que toca
      if (blockStart < rangeStart && isSameDay_(blockStart, rangeStart)) blockStart = new Date(rangeStart); 
      if (blockEnd > rangeEnd && isSameDay_(blockEnd, rangeEnd)) blockEnd = new Date(rangeEnd); 
      let slotStart = new Date(blockStart);

      // analiza en intervaos de media hora el bloque
      while (slotStart < blockEnd && slotStart < rangeEnd) {
        const slotEnd = new Date(slotStart.getTime() + BLOCK_MINUTES * 60 * 1000);
        if (slotEnd > blockEnd || slotEnd > rangeEnd) break;
        if (slotEnd <=  getNow_() && (slotStart = slotEnd)) continue; // puedes asignar una variable como condicion lmaoo

        const busyFixed = fixedBusyEvents.some(ev => {
          // Mira a ver si alguno de los elementos de FixedBusyEvents se solapa con el intervalo de 30 min actual
          return ev.start < slotEnd && ev.end > slotStart;
        });

        if (!busyFixed) {
          // Si no se solapa, es un freeSlot y se intenta meter al array
          const last = mergedFreeSlots[mergedFreeSlots.length - 1]; // como es push siempre miramos el ultimo para comparar
          if (last && last.end.getTime() === slotStart.getTime()) last.end = new Date(slotEnd); 
            // si el anterior intervalo de mergedFreeSlots es consecutivo, en vez de añadirlo, se modifica el anterior
           else mergedFreeSlots.push({start: new Date(slotStart),end: new Date(slotEnd)});
        }
        slotStart = slotEnd;
      }
    });
  }

  Logger.log("freeSlots:");
  Logger.log(JSON.stringify(mergedFreeSlots, null, 2));
  return mergedFreeSlots;
}

// REVISADO
function sortEventsByUrgency_(events) {
    const sortedEvents = [...events];

  // Se ordena directamente el array, ignorando las prioridades dado que luego solo sirven para cosas dentro de sortTowers
  sortedEvents.sort((a, b) => {
    const da = a.deadline ? a.deadline.getTime() : Infinity;
    const db = b.deadline ? b.deadline.getTime() : Infinity;
    if (da !== db) return da - db; // deadline más cercano primero

    // 1 - Prioridad overdue
    if (a.isOverdue || b.isOverdue) {
      if (a.isOverdue && !b.isOverdue) return -1;
      if (!a.isOverdue && b.isOverdue) return 1;

      // Si las dos están pasadas, va primero la que necesita más tiempo
      if (a.etaMinutesRemaining !== b.etaMinutesRemaining) {
        return b.etaMinutesRemaining - a.etaMinutesRemaining;
      }
    }

    // 1.5 - En prioridad 5, las Tareas van antes que Estudio
    if (a.prio === 5 && b.prio === 5 && a.type !== b.type) {
      if (a.type === "Tareas") return -1;
      if (b.type === "Tareas") return 1;
    }

    // 2 - Porcentaje de ocupación del tiempo restante hasta el deadline
    const fillA = getFillRatio_(a);
    const fillB = getFillRatio_(b);
    if (fillA !== fillB) return fillB - fillA;

    return 0;
  });

  Logger.log("sorted thingies");
  Logger.log(JSON.stringify(sortedEvents, null, 2));
  return sortedEvents;
}

/************ GLOBAL STATE ************/

// Contendrá los logs completos leídos desde database (AHORA desde G)
let LOGGED_TASKS = [];


/**
 * 1) Primero hace CUT: mueve (append) todo lo que haya en M (WRITE) hacia G (CHECK),
 *    limpiando la fuente y compactando.
 * 2) Luego lee los logs COMPLETOS desde G (CHECK) y los vuelca en LOGGED_TASKS.
 */
function readLoggedTasks_() {
  const START_ROW = 1497;
  const END_ROW   = 2000;

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("database");
  if (!sheet) {
    throw new Error("No existe la hoja 'database'");
  }

  const srcColLetter  = TASKS_DB_WRITE_COL_LETTER; // "M"
  const destColLetter = TASKS_DB_CHECK_COL_LETTER; // "G"

  // NUEVO: reset del contador por ejecución
  let movedTotal = 0;

  const srcRangeA1 = `${srcColLetter}${START_ROW}:${srcColLetter}${END_ROW}`;
  const srcValues = sheet.getRange(srcRangeA1).getDisplayValues().flat();

  const cleanedSrc = srcValues
    .map(v => (v == null ? "" : String(v).trim()))
    .filter(v => v !== "");

  if (cleanedSrc.length > 0) {
    const destRangeA1 = `${destColLetter}${START_ROW}:${destColLetter}${END_ROW}`;
    const destVals = sheet.getRange(destRangeA1).getDisplayValues().flat();

    let lastNonEmptyRel = -1;
    for (let i = destVals.length - 1; i >= 0; i--) {
      const v = destVals[i];
      if (v != null && String(v).trim() !== "") {
        lastNonEmptyRel = i;
        break;
      }
    }

    let writeRow = START_ROW + (lastNonEmptyRel + 1);
    const capacity = (END_ROW - writeRow + 1);
    if (capacity <= 0) {
      Logger.log("readLoggedTasks_: WARNING -> No hay espacio en " + destRangeA1 + " para mover tareas");
    } else {
      const toMove = cleanedSrc.slice(0, capacity);
      if (toMove.length > 0) {

        // NUEVO: acumula lo movido
        movedTotal += toMove.length;

        const out = toMove.map(x => [x]);
        sheet.getRange(writeRow, columnLetterToIndex_(destColLetter), out.length, 1).setValues(out);

        Logger.log("readLoggedTasks_: CUT moved " + toMove.length + " logs de " + srcColLetter + " -> " + destColLetter +
                   " (" + destColLetter + writeRow + ":" + destColLetter + (writeRow + toMove.length - 1) + ")");

        sheet.getRange(START_ROW, columnLetterToIndex_(srcColLetter), toMove.length, 1).clearContent();
        compactColumnUpInRange_(sheet, srcColLetter, START_ROW, END_ROW);
      }
    }
  }

  // NUEVO: expone el resultado del CUT
  LAST_TASKS_WRITE_MOVED = movedTotal;

  // ============================================================
  // 2) READ: leer AHORA desde CHECK_COL (G) y set global
  //    NEW: ignore "programTime:" line when exporting to LOGGED_TASKS
  // ============================================================
  const rangeA1 = `${destColLetter}${START_ROW}:${destColLetter}${END_ROW}`;
  const values = sheet.getRange(rangeA1).getDisplayValues().flat();

  const cleaned = values
    .map(v => (v == null ? "" : String(v).trim()))
    .filter(v => v !== "")
    // ✅ NUEVO: normaliza cada log quitando el campo programTime (no afecta al resto del texto)
    .map(text => {
      const lines = String(text).split("\n");
      const kept = [];
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (line && line.startsWith("programTime: ")) continue; // IGNORAR este nuevo parámetro
        kept.push(line);
      }
      return kept.join("\n").trim();
    })
    .filter(v => v !== "");

  LOGGED_TASKS = cleaned;

  Logger.log("readLoggedTasks_: READ -> " + cleaned.length + " task logs (rango " + rangeA1 + ")");
  if (cleaned.length > 0) {
    Logger.log("readLoggedTasks_: ejemplo -> " + cleaned[0].slice(0, 200));
  }
}


// NUEVO: cuántos logs se han movido en el último CUT desde I/J/K/L -> A/B/C/D
let LAST_TASKS_WRITE_MOVED = 0;

/**
 * Compacta una columna dentro de un rango [startRow..endRow]:
 * - elimina huecos (celdas vacías intermedias)
 * - mantiene el orden relativo de los no vacíos
 * - NO escribe fuera del intervalo
 */
function compactColumnUpInRange_(sheet, colLetter, startRow, endRow) {
  const colIndex = columnLetterToIndex_(colLetter);
  const numRows = endRow - startRow + 1;

  const range = sheet.getRange(startRow, colIndex, numRows, 1);
  const values = range.getDisplayValues().flat();

  const kept = values
    .map(v => (v == null ? "" : String(v).trim()))
    .filter(v => v !== "");

  const out = [];
  for (let i = 0; i < numRows; i++) {
    out.push([ i < kept.length ? kept[i] : "" ]);
  }

  range.setValues(out);
}


/**
 * Convierte letra(s) de columna (A, B, Z, AA, AB...) a índice 1-based.
 */
function columnLetterToIndex_(letters) {
  let n = 0;
  const s = String(letters).toUpperCase().trim();
  for (let i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n;
}


// REVISADO
function getPendingTasks_(now) {
  // 1) CUT + READ desde database (M -> G, luego lee G)
  try {
    readLoggedTasks_();
  } catch (e) {
    Logger.log("getPendingTasks_: no se pudo cargar LOGGED_TASKS: " + e);
    return [];
  }

  // 2) Parser del formato:
  // "[TASK]\nconcepto: ...\nrowIndex: ...\ndeadline: ...\nprio: ...\netaHours: ...\nprogresoPct: ..."
  function extractField_(logStr, field) {
    if (logStr == null) return null;
    const s = String(logStr);

    // 1) Multilínea: field: value
    let m = s.match(new RegExp("^\\s*" + field + "\\s*:\\s*(.+)\\s*$", "im"));
    if (m && m[1]) return m[1].trim();

    // 2) JSON: "field": "..."
    m = s.match(new RegExp('"' + field + '"\\s*:\\s*"([^"]+)"', "i"));
    if (m && m[1]) return m[1].trim();

    // 3) field=...
    m = s.match(new RegExp("(^|\\s|[,{])" + field + "\\s*=\\s*([^\\n\\r,}\\]]+)", "i"));
    if (m && m[2]) return m[2].trim();

    return null;
  }

  function parseTaskFromLog_(logStr) {
    const conceptoRaw = extractField_(logStr, "concepto");
    if (!conceptoRaw) return null;

    const rowIndexRaw = extractField_(logStr, "rowIndex");
    const deadlineRaw = extractField_(logStr, "deadline");
    const prioRaw     = extractField_(logStr, "prio");
    const etaRaw      = extractField_(logStr, "etaHours");
    const progRaw     = extractField_(logStr, "progresoPct");

    const concepto = String(conceptoRaw).trim();
    const rowIndex = Number(rowIndexRaw) || null;

    let deadline = null;
    if (deadlineRaw != null && String(deadlineRaw).trim() !== "" && String(deadlineRaw).toLowerCase() !== "null") {
      const parsed = new Date(Date.parse(String(deadlineRaw)));
      if (!isNaN(parsed.getTime())) deadline = parsed;
    }

    // ajustamos la deadline para que sea el dia de antes a las 23:59
    if (deadline instanceof Date && !isNaN(deadline)) {
      const effDeadline = new Date(deadline.getTime());
      effDeadline.setDate(effDeadline.getDate());
      effDeadline.setHours(23, 59, 0, 0);
      deadline = effDeadline;
    }

    let prio = Number(prioRaw) || 0;
    const etaHours = Number(etaRaw) || 0;

    // progresoPct puede venir vacío; por defecto 0
    let progreso = 0;
    if (progRaw != null && String(progRaw).trim() !== "") {
      progreso = (Number(progRaw) || 0) / 100;
    }

    if (!concepto || !etaHours) return null;

    return { concepto, rowIndex, deadline, prio, etaHours, progreso };
  }

  const tasks = [];

  for (let i = 0; i < (LOGGED_TASKS || []).length; i++) {
    const parsed = parseTaskFromLog_(LOGGED_TASKS[i]);
    if (!parsed) continue;

    const { concepto, rowIndex, deadline, prio: prioIn, etaHours, progreso } = parsed;

    const rawMinutes = etaHours * 60 * (1 - progreso);
    if (rawMinutes <= 0) continue; // para que ignore las que están completas (progreso >= 100)

    let remainingMinutes = rawMinutes;

    let etaMinutesRemaining = Math.ceil(remainingMinutes / 60) * 60; //lo redondeamos a la hora superior para que sea más facil colocarlo
    if (etaMinutesRemaining <= 0) continue;

    let prio = prioIn;
    const isOverdue = (deadline !== null && deadline.getTime() < now.getTime()); // booleano de si ya ha pasado la fecha
    if (isOverdue) prio = 5;

    // start = ahora
    const start = now;

    // allowedDays = todos los días desde ahora hasta deadline
    // si no hay deadline, lo limitamos al día de hoy (evita loops raros)
    const allowedDays = [];
    if (deadline === null) {
      allowedDays.push(dateKeyFromDate_(startOfDay_(start)));
    } else {
      let d = startOfDay_(start);
      const endD = startOfDay_(deadline);
      while (d <= endD) {
        allowedDays.push(dateKeyFromDate_(d));
        d = addDays_(d, 1);
      }
    }

    tasks.push({
      type: "Tareas",
      rowIndex: (rowIndex == null ? -1 : rowIndex),
      concepto,
      start,
      deadline,
      prio,
      bracket: prio,                   // bracket = prio
      allowedDays,
      etaMinutesRemaining,
      totalMinutesOriginal: etaMinutesRemaining,
      blocksTotal: etaMinutesRemaining / 30,
      isOverdue
    });
  }
  Logger.log("Tareas:");
  Logger.log(JSON.stringify(tasks,null,2));
  return tasks;
}


function buildProjectPriorityArray_(now, horizon) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("database");
  if (!sheet) return [];

  const range = sheet.getRange("E5:I12"); // E..I
  const values = range.getValues();

  const result = [];
  let seq = 1;

  for (let r = 0; r < values.length; r++) {
    const row = values[r];

    const projectName = (row[0] || "").toString().trim(); // Col E
    const weeklyHoursRaw = row[4];                        // Col I

    if (!projectName) continue;

    const weeklyHours = Number(weeklyHoursRaw);
    if (!isFinite(weeklyHours) || weeklyHours <= 0) continue;

    // Construimos intervalos [start,end) hacia atrás desde horizon
    let cursorEnd = new Date(horizon);

    // Si horizon <= now, no hay nada que planificar
    if (cursorEnd.getTime() <= now.getTime()) continue;

    while (cursorEnd.getTime() > now.getTime()) {
      let cursorStart = addDays_(cursorEnd, -7);

      // Recorte del último intervalo para que empiece en el momento actual
      if (cursorStart.getTime() < now.getTime()) {
        cursorStart = new Date(now);
      }

      // Horas escaladas proporcionalmente a la duración del intervalo
      const scaledHours = getScaledProjectHours_(weeklyHours, cursorStart, cursorEnd);

      // Si al escalar se queda en 0, no generamos evento
      if (scaledHours > 0) {
        const minutes = scaledHours * 60;

        const allowedDays = buildAllowedDaysKeys_(cursorStart, cursorEnd);

        // Deadline a las 23:59 del día de cursorEnd (en vez de la hora que traiga cursorEnd)
        const deadlineDay = startOfDay_(new Date(cursorEnd));
        const deadline2359 = new Date(
          deadlineDay.getFullYear(),
          deadlineDay.getMonth(),
          deadlineDay.getDate(),
          23, 59, 0, 0
        );

        result.push({
          bracket: 1,
          concepto: projectName,
          type: "Proyecto",
          etaMinutesRemaining: minutes,
          id: "P#" + (seq++),
          isOverdue: false,
          prio: 1,
          totalMinutesOriginal: minutes,
          blocksTotal: Math.ceil(minutes / BLOCK_MINUTES),
          allowedDays: allowedDays,
          // La "deadline" del intervalo es el final del chunk, pero fijada a 23:59 del día
          deadline: deadline2359,
          // Start informativo del intervalo (cursorStart)
          start: new Date(cursorStart)
        });
      }

      cursorEnd = new Date(cursorStart);
    }
  }

  Logger.log("proyectos:");
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}



/**
 * Escala horas semanales (baseHours) proporcionalmente a la duración del intervalo.
 * - 7 días completos => baseHours completas
 * - Menos de 7 días => baseHours * (duración / 7 días)
 * - Redondeo hacia arriba a múltiplos de 30 min, devolviendo horas (.0 / .5)
 * - Aplica el mismo cap que tu prio===1 en getScaledStudyHours_
 */
function getScaledProjectHours_(baseHours, start, end) {
  if (!start || !end || baseHours <= 0) return 0;

  const totalWeekMs = 7 * 24 * 60 * 60 * 1000;
  const durMs = Math.max(0, end.getTime() - start.getTime());
  if (durMs <= 0) return 0;

  const scaled = baseHours * (durMs / totalWeekMs);

  // Redondeo hacia arriba cada media hora (en horas)
  let roundedHours = Math.ceil((scaled * 60) / 30) * 30 / 60;

  // Mínimo 60 minutos (1 hora) si hay solape positivo
  if (roundedHours > 0 && roundedHours < 1) {
    roundedHours = 1;
  }

  return roundedHours;
}



/**
 * Devuelve allowedDays como array de dayKeys ("yyyy-MM-dd") entre start y end.
 * Nota: end se trata como límite superior; se incluye el día de start y los días intermedios.
 */
function buildAllowedDaysKeys_(start, end) {
  const tz = Session.getScriptTimeZone();

  const days = [];
  let cursor = startOfDay_(new Date(start));
  const endDay = startOfDay_(new Date(end));

  // Incluimos días hasta el día anterior a end si end cae justo en startOfDay del siguiente
  // En la práctica, para tu scheduling va bien incluir el día de end también si hay tiempo ese día.
  while (cursor.getTime() <= endDay.getTime()) {
    days.push(Utilities.formatDate(cursor, tz, "yyyy-MM-dd"));
    cursor = addDays_(cursor, 1);
  }

  return days;
}



// REVISADO
function getScaledStudyHours_(priority, start, end, now) {
  const baseHours = BASE_STUDY_HOURS[priority] || 0;
  if (!start || !end || baseHours <= 0) return 0;

  const weekStart = new Date(now), weekEnd = addDays_(weekStart, 7); // intervalo para la semana
  const overlapStart = Math.max(start.getTime(), weekStart.getTime()); // La fecha posterior de las dos es el inicio
  const overlapEnd = Math.min(end.getTime(), weekEnd.getTime()); // La fecha anterior de las dos es la final
  const overlapMs = Math.max(0, overlapEnd - overlapStart); // Restamos las dos para que salga un numero en ms
  const totalWeekMs = weekEnd - weekStart;

  if(priority === 1){
    return Math.min(Math.ceil((baseHours * 60 * (overlapMs / totalWeekMs)) / 30) * 30 / 60,180);
  }
  // Redondea hacia arriba cada media hora pero te lo saca en horas (saca o .0 o .5)
  return Math.ceil((baseHours * 60 * (overlapMs / totalWeekMs)) / 30) * 30 / 60;
}

// REVISADO
function buildExamPriorityArray_() {

  // DATOS PARA OPERAR
  // ANTES: leía directamente del Calendar (cal.getEvents) -> ahora usa LOGGED_CALENDAR_EVENTS["Examenes"]
  const now = getNow_();
  const today = startOfDay_(now);
  const windowEnd = getDynamicHorizonEnd_(now);

  // --- NEW: cargar multiplicadores Estudio desde 'database'!J5:M12 (nombre en J, factor en M) ---
  // Estructura esperada: col J = nombre, col M = multiplicador (number)
  const multSheet = SpreadsheetApp.getActive().getSheetByName(DATABASE_SHEET_NAME);
  let ESTUDIO_MULTIPLIERS = {}; // { "nombrelower": factor }
  if (multSheet) {
    try {
      const multRange = multSheet.getRange("J5:M12").getValues(); // [row][0..3] => J..M
      for (let r = 0; r < multRange.length; r++) {
        const name = multRange[r][0]; // J
        const factor = multRange[r][3]; // M
        const nameKey = String(name || "").trim().toLowerCase();
        const fNum = Number(factor);
        if (nameKey && isFinite(fNum) && fNum > 0) {
          ESTUDIO_MULTIPLIERS[nameKey] = fNum;
        }
      }
    } catch (e) {
      Logger.log("buildExamPriorityArray_: leyendo multiplicadores J5:M12 falló: " + e);
    }
  } else {
    Logger.log("buildExamPriorityArray_: no se encontró la sheet '" + DATABASE_SHEET_NAME + "' para multiplicadores.");
  }

  // Devuelve el multiplicador para un concepto:
  // 1) match exacto por texto
  // 2) si no, match por substring (elige el key más largo que aparezca en el concepto)
  function getMultiplierForConcept_(concepto) {
    const c = String(concepto || "").trim().toLowerCase();
    if (!c) return 1;

    if (ESTUDIO_MULTIPLIERS[c] != null) return ESTUDIO_MULTIPLIERS[c];

    let bestKey = "";
    let bestFactor = 1;

    const keys = Object.keys(ESTUDIO_MULTIPLIERS);
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      if (k && c.indexOf(k) !== -1) {
        if (k.length > bestKey.length) {
          bestKey = k;
          bestFactor = ESTUDIO_MULTIPLIERS[k];
        }
      }
    }
    return bestFactor || 1;
  }

  // Aplica multiplicador a minutos (manteniendo bloques de 30m) y mínimo 60m por intervalo
  function applyMultiplierWithMinimum_(concepto, minutes) {
    const base = Number(minutes) || 0;
    if (base <= 0) return 0;

    const factor = getMultiplierForConcept_(concepto);
    const scaled = base * factor;

    // redondeo a bloques de 30 min (hacia arriba, para no perder tiempo al multiplicar)
    const rounded = Math.ceil(scaled / 30) * 30;

    // mínimo 60 minutos por intervalo
    return Math.max(60, rounded);
  }

  // Aseguramos que el global esté cargado (por si alguien llama esto sin haber hecho READ antes)
  if (
    !LOGGED_CALENDAR_EVENTS ||
    !LOGGED_CALENDAR_EVENTS["Examenes"] ||
    LOGGED_CALENDAR_EVENTS["Examenes"].length === 0
  ) {
    try {
      readLoggedCalendarEvents_();
    } catch (e) {
      Logger.log("buildExamPriorityArray_: readLoggedCalendarEvents_ falló: " + e);
    }
  }

  const examLogs = (LOGGED_CALENDAR_EVENTS && LOGGED_CALENDAR_EVENTS["Examenes"]) ? LOGGED_CALENDAR_EVENTS["Examenes"] : [];

  // Parser simple del formato:
  // "[GOOGLE CALENDAR EVENT]\ntitle: ...\n...\nstart: Fri Jan 02 2026 11:00:00 GMT+0100 ...\nend: ..."
  function parseExamLog_(logText) {
    if (!logText) return null;

    const text = String(logText);

    const mTitle = text.match(/^\s*title:\s*(.*)\s*$/mi);
    const mStart = text.match(/^\s*start:\s*(.*)\s*$/mi);
    const mEnd   = text.match(/^\s*end:\s*(.*)\s*$/mi);

    const titleRaw = mTitle ? String(mTitle[1]).trim() : "";
    const startRaw = mStart ? String(mStart[1]).trim() : "";
    const endRaw   = mEnd   ? String(mEnd[1]).trim()   : "";

    if (!titleRaw || !startRaw) return null;

    const start = new Date(startRaw);
    const end = endRaw ? new Date(endRaw) : null;

    if (!(start instanceof Date) || isNaN(start)) return null;
    if (endRaw && (!(end instanceof Date) || isNaN(end))) return null;

    return {
      title: titleRaw,
      start: start,
      end: end
    };
  }

  const parsedExams = [];
  for (let i = 0; i < examLogs.length; i++) {
    const parsed = parseExamLog_(examLogs[i]);
    if (!parsed) continue;

    // Fuera de ventana (por fecha del examen)
    const examDate = startOfDay_(parsed.start);
    if (examDate < today || examDate > windowEnd) continue;

    parsedExams.push(parsed);
  }

  Logger.log("Found %s exam logs in window (%s total raw logs in database).", parsedExams.length, examLogs.length);

  const result = [];

  // Intentamos colocar los intervalos de prioridad para todos los examenes en las proximas 4 semanas
  parsedExams.forEach(exam => {
    const title = exam.title;
    const examStart = exam.start;
    const examDate = startOfDay_(examStart);

    // Fuera de ventana
    if (examDate < today || examDate > windowEnd) return;

    // Vamos en chunks de 1 semana desde "now" hasta el examen.
    // Cada chunk genera un evento de estudio independiente.
    let cursorStart = new Date(now);
    cursorStart = new Date(Math.max(cursorStart.getTime(), today.getTime()));

    while (cursorStart < examStart) {

      // Fin del intervalo semanal
      let cursorEnd = addDays_(startOfDay_(cursorStart), 7);
      cursorEnd.setHours(
        examStart.getHours(),
        examStart.getMinutes(),
        0,
        0
      );

      // Si nos pasamos del examen, recortamos
      if (cursorEnd > examStart) {
        cursorEnd = new Date(examStart);
      }

      // Días restantes desde el INICIO del chunk
      const diffDaysChunk = Math.floor(
        (examDate - startOfDay_(cursorStart)) / 86400000
      );

      // Cálculo de prioridad
      let prio;
      if (diffDaysChunk >= 28) {
        prio = 1; // se puede repetir indefinidamente
      } else {
        // 0–6 → 5, 7–13 → 4, 14–20 → 3, 21–27 → 2
        prio = 5 - Math.floor(diffDaysChunk / 7);
        prio = Math.min(5, Math.max(2, prio));
      }

      const baseHours = BASE_STUDY_HOURS[prio] || 0;

      if (baseHours > 0) {
        const intervalMs = Math.max(
          0,
          cursorEnd.getTime() - cursorStart.getTime()
        );
        const weekMs = 7 * 86400000;

        // Escalado proporcional si no es semana completa
        let scaledHours =
          Math.ceil((baseHours * (intervalMs / weekMs) * 60) / 30) *
          30 /
          60;
        if (prio === 1) scaledHours = Math.min(scaledHours,3);
        let minutes = Math.ceil((scaledHours * 60) / 30) * 30;

        // --- NEW: aplicar multiplicador por concepto + mínimo 60m ---
        minutes = applyMultiplierWithMinimum_(title, minutes);

        if (minutes > 0) {
          // allowedDays del intervalo
          const allowedDays = [];
          let d = startOfDay_(cursorStart);
          const endD = startOfDay_(cursorEnd);

          while (d <= endD) {
            allowedDays.push(dateKeyFromDate_(d));
            d = addDays_(d, 1);
          }

          result.push({
            type: "Estudio",
            start: new Date(cursorStart),
            deadline: new Date(cursorEnd),
            concepto: title,
            prio: prio,
            bracket: prio,
            isOverdue: false,
            allowedDays: allowedDays,
            etaMinutesRemaining: minutes,
            totalMinutesOriginal: minutes,
            blocksTotal: minutes / 30
          });
        }
      }

      // Avanzamos al siguiente chunk semanal
      cursorStart = new Date(cursorEnd);
    }
  });

  // Una vez que hemos colocado todos los examenes, miramos a ver si alguno de las asignaturas está sin examen
  
  // Hacemos un set con todas las asignaturas
  const subjectsWithoutExams = new Set(SUBJECTS);
  // Si algun examen está en result, lo eliminamos del set
  result.forEach(item  => {
    const conceptoLower = String(item.concepto || "").toLowerCase(); // lowercase porsiaca
    SUBJECTS.forEach(subj => {
      const subjLower = subj.toLowerCase();
      if (conceptoLower.indexOf(subjLower) !== -1) subjectsWithoutExams.delete(subjLower); //los buscamos en el array de arriba
    });

  });

  // para las asignaturas que sobran les hacemos eventos de prioridad 1
  // PERO en intervalos semanales repetidos desde ahora hasta el final del horizonte
  subjectsWithoutExams.forEach(subj => {
    const priority = 1;

    let cursorStart = new Date(now); // desde el momento de ejecución
    cursorStart = new Date(Math.max(cursorStart.getTime(), today.getTime()));

    while (cursorStart < windowEnd) {
      // fin del chunk semanal
      let cursorEnd = addDays_(startOfDay_(cursorStart), 7);

      // recortamos si nos pasamos del horizonte
      if (cursorEnd > windowEnd) cursorEnd = new Date(windowEnd);

      // horas objetivo de esa semana (escala si no es semana completa)
      const baseHours = BASE_STUDY_HOURS[priority] || 0;
      if (baseHours > 0) {
        const intervalMs = Math.max(0, cursorEnd.getTime() - cursorStart.getTime());
        const weekMs = 7 * 86400000;

        const rawMinutes = baseHours * (intervalMs / weekMs) * 60;
        let minutes = Math.min(Math.floor((rawMinutes - 1e-9) / 30) * 30,180);

        // --- NEW: aplicar multiplicador por concepto + mínimo 60m ---
        minutes = applyMultiplierWithMinimum_(subj, minutes);

        // mantenemos el tope histórico de este bloque (si lo quieres eliminar, dímelo)
        minutes = Math.min(minutes, 180);

        if (minutes > 0) {
          // allowedDays del chunk
          const allowedDays = [];
          let d = startOfDay_(cursorStart);
          const endD = startOfDay_(cursorEnd);
          while (d <= endD) {
            allowedDays.push(dateKeyFromDate_(d));
            d = addDays_(d, 1);
          }

          result.push({
            type: "Estudio",
            start: new Date(cursorStart),
            deadline: new Date(cursorEnd),
            concepto: subj,
            prio: priority,
            bracket: priority,
            isOverdue: false,
            allowedDays: allowedDays,
            etaMinutesRemaining: minutes,
            totalMinutesOriginal: minutes,
            blocksTotal: Math.min(minutes / 30,180)
          });
        }
      }

      // siguiente semana
      cursorStart = new Date(cursorEnd);
    }
  });

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}


// REVISADO
// Esta funcion es la más complicada de todo el código, está separada por secciones y antes de cada seccion hay una explicación, además de los
// comentarios de cada cosa
// Nota: Aun queda por programar la parte que, para días con muchos eventos, purgue el día para que fisicamente quepan
function optimizeScheduleWithTowers_(freeSlots, eventsByBracket) {
  

  Logger.log("lo que entra a las torres");
  Logger.log(JSON.stringify(eventsByBracket, null, 2));
  // 1- PREPARACIÓN DE DATOS INICIALES
  // Extracción de los minutos por día
  const freeMinutesByDay = {};
  let globalRangeStart = null;
  let globalRangeEnd = null;

  freeSlots.forEach(slot => {
    const s = new Date(slot.start), e = new Date(slot.end);

    // Actualizamos los límites globales del rango
    if (!globalRangeStart || s < globalRangeStart) globalRangeStart = s;
    if (!globalRangeEnd   || e > globalRangeEnd)   globalRangeEnd   = e;

    // Vamos sumando minutos al dia en el diccionario
    const minutes = (e.getTime() - s.getTime()) / (60 * 1000);
    const key = dateKeyFromDate_(s);
    freeMinutesByDay[key] = (freeMinutesByDay[key] || 0) + minutes;
  });

  // Creacion de los arrays y dicts que nos servirán para luego
  const dayKeys = Object.keys(freeMinutesByDay).sort();
  const eventsMeta = [];
  const eventMetaById = {};
  let eventCounter = 0;
  
  eventsByBracket.forEach(ev => {
    const id = "E#" + (eventCounter++);
    const meta = { ...ev, id };
    eventsMeta.push(meta);
    eventMetaById[id] = meta;
  });

  // 2- ESTADO INICIAL DEL ALGORITMO
  // Primero, se va a intentar colocar todas las horas al final de su plazo, para crear la máxima disparidad posible
  const dayUsedBlocks = {};
  dayKeys.forEach(k => { dayUsedBlocks[k] = 0; });
  const eventDayBlocks = {};

  eventsMeta.forEach(meta => {
    // hace el bucle para cada evento
    const { id, type, blocksTotal, allowedDays } = meta;
    let remainingBlocks = blocksTotal;
    if (remainingBlocks <= 0) return;

    eventDayBlocks[id] = eventDayBlocks[id] || {concepto:meta["type"] + ": " + meta["concepto"]}; // Lo inicializa si está vacío

    // Mira en que dias puede poner los eventos, pero solo incluye los dias dentro del intervalo actual de freeSlots
    const allowedDesc = [...allowedDays]
      // Probamos con la lista de minutos libres por dia, y la usamos de referencia para que si un dia no está en ese rango, no la cuente,
      // que si no jode el algoritmo
      .filter(k => Object.prototype.hasOwnProperty.call(freeMinutesByDay, k))
      // Además, el día tiene que tener hueco útil antes de la hora de deadline de ese evento
      .filter(k => isDayUsableForEvent_(meta, k))
      .sort((a, b) =>
        dateFromKey_(b).getTime() - dateFromKey_(a).getTime()
      );

    // Si ninguno de los días permitidos está en el horizonte actual, no intentamos
    // colocarlo ahora; se colocará en futuros runs cuando el horizonte avance.
    if (allowedDesc.length === 0) return;

    // luego para cada evento, itera por dias
    for (let i = 0; i < allowedDesc.length && remainingBlocks > 0; i++) {
      const dayKey = allowedDesc[i];
      // Saca cuantos bloques de este evento tengo por ahora en este día
      const currentBlocksUsed = eventDayBlocks[id][dayKey] || 0;
      
      const MAX_BLOCKS_PER_DAY = 6; 
      // Si el bloque se puede colocar entero, el min será remainingBlocks, si no el máximo que se
      let capacityForThisEvent = Math.min(remainingBlocks, MAX_BLOCKS_PER_DAY - currentBlocksUsed);
      if (capacityForThisEvent <= 0) continue;

      let toAssign = capacityForThisEvent; // esto si es tasks se resuelve directamente aqui

      if (type === "Estudio") {
        // Solo vamos a poner bloques de más de una hora en Estudio, este bloque comprueba eso
        if (currentBlocksUsed === 0) {
          if (remainingBlocks === 1) {
            continue;
          }
          toAssign = Math.max(toAssign, 2); // Inicio minimo 1h
        }
        toAssign = Math.min(toAssign, capacityForThisEvent);

        const finalCount = currentBlocksUsed + toAssign;
        if (finalCount === 1) continue; //  si sobra uno, no se va a colocar en un día
      }

      if (toAssign <= 0) continue; // si es 0 se salta directamente

      eventDayBlocks[id][dayKey] = currentBlocksUsed + toAssign; // actualizamos cuanto se ha usado del evento
      dayUsedBlocks[dayKey] = (dayUsedBlocks[dayKey] || 0) + toAssign; // actualizamos cuanto se ha usado del dia
      remainingBlocks -= toAssign; // para el siguiente bucle
    }

    if (remainingBlocks > 0)  Logger.log("Evento '" + (meta.concepto || meta.id) +  "' " + "con id: " +  meta.deadline + " y prio " + meta.prio + " no ha podido colocar " + (remainingBlocks * 30) + " minutos en la fase base.");
  });


  // Resultados del 1er bucle
  let stats = computeOverflowStats_(dayUsedBlocks);
  Logger.log("Estado inicial (tras fase 1, antes de rebalancear torres):");
  Logger.log(JSON.stringify({
    freeMinutesByDay: freeMinutesByDay,
    dayUsedBlocks: dayUsedBlocks,
    stats: stats
  }, null, 2));

  // 3- ITERACION POR PAREJAS
  // Vamos a probar pareja por pareja (empezando por la más alta como source y pillando targets) a ver si, al mover un bloque de intervalo de sitio,
  // la diferencia entre el overflow de las parejas y la media baja, iterando primero desde arriba (iteraciones generales), luego una iteracion
  // por parejas que se reiniciará cada vez que un emparejamiento produzca una mejora, y luego dentro de cada emparejamiento, una iteración de todos los eventos
  // del source para ver cual es el mejor movimiento que se puede realizar
  let improvedGlobal = false;
  const MAX_GLOBAL_ITER = 5000;
  const MAX_PAIR_STEPS = 100;
  for (let iter = 0; iter < MAX_GLOBAL_ITER; iter++) {
    stats = computeOverflowStats_(dayUsedBlocks);
    const overflow = stats.overflow;

    // Inicializamos el dia con su iverflow
    const dayInfos = dayKeys.map(dayKey => ({
      dayKey: dayKey,
      ov: overflow[dayKey] || 0,
      usedBlocks: dayUsedBlocks[dayKey] || 0
    }));

    if (dayInfos.length < 2) break;

    // Ordenamos los dias por overflow
    dayInfos.sort((a, b) => a.ov - b.ov);

    const n = dayInfos.length;
    let anyImprovementThisRound = false;

    // Recorremos todas las posibles parejas (high, low)
    // si hay alguna mejora en alguna iteracion, anyImprovementThisRound va a salir del for loop y empezar otra vez desde 
    // el for loop de MAX_GLOBAL_ITER (el de justo encima) 
    for (let hiIdx = n - 1; hiIdx >= 0 && !anyImprovementThisRound; hiIdx--) {
      // Cogemos aquí la información del dia High
      const highInfo = dayInfos[hiIdx];
      const sourceDay = highInfo.dayKey;

      // si no hay bloques en el "high", no se puede mover nada desde él
      if ((dayUsedBlocks[sourceDay] || 0) <= 0) continue;

      for (let loIdx = 0; loIdx < n && !anyImprovementThisRound; loIdx++) {
        if (loIdx === hiIdx) continue; //para que no se empareje consigo mismo

        // Información de los dias que vamos a comprar
        const lowInfo = dayInfos[loIdx];
        const targetDay = lowInfo.dayKey;
        const ovHigh = highInfo.ov;
        const ovLow = lowInfo.ov;
        if (ovHigh <= ovLow) continue; // Si son iguales no vas a mejorar nada asiq lo dejas


        // COMPRACIÓN POR PAREJAS
        let improvedAny = false;
        for (let steps = 0; steps < MAX_PAIR_STEPS; steps++) {
          stats = computeOverflowStats_(dayUsedBlocks);
          const overflowPair = stats.overflow; // De donde sacas el oveflow

          // Miramos los overflows de los dos dias
          const ovSource = overflowPair[sourceDay] || 0; 
          const ovTarget = overflowPair[targetDay] || 0;

          // Si alguno de los target es más grande que el source, ya no es el más con diferencia y vamos a otro
          if (ovSource <= ovTarget) break;

          // Parametros de evaluacion
          const avg = (ovSource + ovTarget) / 2; // Constante 
          const currentMetric = Math.abs(ovSource - avg) + Math.abs(ovTarget - avg);

          let bestMove = null;
          let bestMetric = currentMetric;
          // aqui iteramos los eventos para encontrar el mejor movimiento posible
          for (const evId in eventDayBlocks) {
            // sacamos los datos del evento
            const meta = eventMetaById[evId];
            if (!meta) continue;

            const allowed = meta.allowedDays;
            if (allowed.indexOf(sourceDay) === -1 || allowed.indexOf(targetDay) === -1) continue; // si el evento no se puede en ninguno lo saltamos

            // El movimiento solo tiene sentido si el día es físicamente usable para ese evento (por hora de deadline)
            if (!isDayUsableForEvent_(meta, sourceDay) || !isDayUsableForEvent_(meta, targetDay)) continue;
            
            // Datos para probar el movimiento
            const type = meta.type;
            const blocksOnSource = eventDayBlocks[evId][sourceDay] || 0;
            const blocksOnTarget = eventDayBlocks[evId][targetDay] || 0;
            if (blocksOnSource <= 0) continue; // si en el source no hay nada de ese evento, no podemos mover cosas y lo saltamos

            // Para probar movimientos, como Estudio tiene una duración minima de 1h, puede que en target no haya de ese
            // evento y no se pueda mover solo media hora, asi que permitimos los dos movimientos
            const stepOptions = (type === "Estudio" || type === "Proyecto" ) ? [1, 2] : [1];

            // Probamos a encontrar un movimiento para ese evento que mejore la situación
            for (let s = 0; s < stepOptions.length; s++) {
              const stepBlocks = stepOptions[s]; // el movimiento que estamos probando
              if (blocksOnSource < stepBlocks) continue;

              // Evaluamos los nuevos conteos de ese evento en ambos sitios
              const newSourceCount = blocksOnSource - stepBlocks;
              const newTargetCount = blocksOnTarget + stepBlocks;


              if (type === "Estudio" || type === "Proyecto" )  {
                // esta arrow function dice si el movimiento es valido en base a si en ambos se queda un conteo valido de intervalos de media hora
                const validEstudio = c => c === 0 || (c >= 2 && c <= 6); 
                if (!validEstudio(newSourceCount) || !validEstudio(newTargetCount)) continue;
              } else {
                const validTarea = c => c >= 0 && c <= 6; // igual que arriba solo que para tasks
                if (!validTarea(newSourceCount) || !validTarea(newTargetCount)) continue;
              }

              const deltaMinutes = stepBlocks * 30;
              const newOvSource = ovSource - deltaMinutes;
              const newOvTarget = ovTarget + deltaMinutes;

              // Si el movimiento solo permuta los valores, es un movimiento ilegal y nos lo saltamos
              if (newOvSource === ovTarget && newOvTarget === ovSource) continue;

              // probamos si la diferencia respecto a la media de ambos ha cambiado, y si ha mejorado, lo metemos en bestMove
              const newMetric = Math.abs(newOvSource - avg) + Math.abs(newOvTarget - avg);
              if (newMetric + 1e-6 < bestMetric) {
                bestMetric = newMetric;
                bestMove = {evId: evId, sourceDay: sourceDay, targetDay: targetDay,
                            stepBlocks: stepBlocks, newOvSource: newOvSource, newOvTarget: newOvTarget};
                }
            }
          }

          if (!bestMove) break; // en caso de que no haya mejor movimiento, se pasa a la siguiente pareja

          // REALIZACIÓN DEL MOVIMIENTO DE EVENTOS ENTRE LAS PAREJAS
          // sacamos todos los valores de bestMove a variables individuales (destructuring)
          const {
            evId,
            sourceDay: sDay,
            targetDay: tDay,
            stepBlocks,
            newOvSource,
            newOvTarget
          } = bestMove;
          const metaMove = eventMetaById[evId]; //sacamos el evento al que se refiere el movimiento

          const prevSourceCount = eventDayBlocks[evId][sDay] || 0;
          const prevTargetCount = eventDayBlocks[evId][tDay] || 0;

          // quitamos ese numero de bloques del dia donde origina el movimiento (source o sDay)
          eventDayBlocks[evId][sDay] = prevSourceCount - stepBlocks;
          // si se queda vacío el día, eliminamos ese dia de los dias donde está ese evento
          if (eventDayBlocks[evId][sDay] <= 0) delete eventDayBlocks[evId][sDay];
          // metemos esos dias en el target (tday)
          eventDayBlocks[evId][tDay] = prevTargetCount + stepBlocks; 

          // modificamos los conteos de bloques usados de tDay y sDay
          dayUsedBlocks[sDay] = (dayUsedBlocks[sDay] || 0) - stepBlocks;
          dayUsedBlocks[tDay] = (dayUsedBlocks[tDay] || 0) + stepBlocks;

          Logger.log(
            "Movimiento parejas: " + (stepBlocks * 30) + " min del evento '" +
            (metaMove.concepto || evId) + "' de " + sDay + " a " + tDay +
            " (overflow " + ovSource + "," + ovTarget + " → " +
            newOvSource + "," + newOvTarget + ")"
          );

          improvedAny = true; // para ver que se ha mejorado algo
        }
        // en caso de que el emparejamiento haya mejorado algo, significa que aun hay cosas que mejorar y se sigue
        if (improvedAny) {
          anyImprovementThisRound = true; // si esto es true, se sale del for loop de parejas y se empieza de cero
          improvedGlobal = true;
        }
      }
    }
    // si al realizar una iteracion de todas las parejas no se ha encontrado ninguna mejora, significa que está en el estado optimo y se sale
    if (!anyImprovementThisRound) break;
  }

  // sacamos los resultados si algo ha mejorado (por lo general va a siempre mejorar)
  if (improvedGlobal) {
    stats = computeOverflowStats_(dayUsedBlocks);
    Logger.log("Estado tras rebalanceo por parejas (torres):");
    Logger.log(JSON.stringify({
      dayUsedBlocks: dayUsedBlocks,
      stats: stats
    }, null, 2));
  }

  
  // Recalcular stats tras meter el extra
  stats = computeOverflowStats_(dayUsedBlocks);


  const resultPre = {
    freeMinutesByDay: freeMinutesByDay,
    finalDayUsedMinutes: Object.fromEntries(
      dayKeys.map(k => [k, (dayUsedBlocks[k] || 0) * 30])
    ),
    overflowFinal: stats.overflow,
    meanOverflowFinal: stats.meanOverflow,
    FFinal: stats.F,
    eventDayBlocks: eventDayBlocks,
    eventsMeta: eventsMeta,
    dayKeys: dayKeys
  };

  Logger.log("Resultado pre purga optimización por torres (precalendario lógico):");
  Logger.log(JSON.stringify(resultPre, null, 2));


  // FASE DE EMPEQUEÑECIMIENTO DE DIAS LLENOS
  // Una vez que hemos encontrado la distribución optima, algunos dias puede que necesiten más horas que las que hay disponibles, asi que toca filtrar
  // La prioridad será siempre Tasks 5 > Estudio 5 >  Resto de osas

  // Copiamos el overflow para ver que hay cada cosa
  let overflowNow = stats.overflow || {};

  const extraTimes = {};// para los dias que necesito meter tiempo extra

  // iteramos cada dia por separado probando a ver cuales tienen overflow positivo
  dayKeys.forEach(dayKey => {
    let ovMinutes = overflowNow[dayKey] || 0; // los overflows del dia
    if (ovMinutes <= 0) return; 
    // Si pasa esto es que toca purgar
    Logger.log("estoy dentro de un dia de purgado");

    let blocksToRemove = Math.ceil(ovMinutes / 30); // El numero de bloques a quitar
    Logger.log("bloques a eliminar: " + blocksToRemove);
    Logger.log("eventDayBlocks: ");
    Logger.log(JSON.stringify(eventDayBlocks, null, 2));
    

    // Construimos la lista de candidatos de ese día que pueden perder bloques
    const candidates = [];

    for (const evId in eventDayBlocks) {
      const meta = eventMetaById[evId];
      if (!meta) continue;

      const blocks = eventDayBlocks[evId][dayKey] || 0;
      if (blocks <= 0) continue;

      const isStudy = (meta.type === "Estudio");
      const minBlocks = isStudy ? 4 : 2; // El estudio (a primeras) no puede bajar d 2h, el resto no puede bajar de 1h

      if (blocks > minBlocks) {
        candidates.push({
          concepto: eventDayBlocks[evId]["concepto"],
          evId: evId,
          blocks: blocks,
          isStudy: isStudy
        });
      }
    }

    Logger.log("candidatos");
    Logger.log(candidates);

    while (candidates.length > 0 && blocksToRemove > 0){
      candidates.sort((a, b) => b.blocks - a.blocks);
      for (let i = 0; i < candidates.length && candidates.length > 0; i++) {
        const evId  = candidates[i]["evId"];
        const meta = eventMetaById[evId];
        const isStudy = (meta.type === "Estudio");
        const minBlocks = isStudy ? 4 : 2;
        const currentBlocks = eventDayBlocks[evId][dayKey] || 0;
        eventDayBlocks[evId][dayKey] = currentBlocks - 1;
        dayUsedBlocks[dayKey] = (dayUsedBlocks[dayKey] || 0) - 1;
        Logger.log("bloque purgado: "+ eventDayBlocks[evId]["concepto"]);
        Logger.log("Bloques antes: " + blocksToRemove);
        blocksToRemove--;
        Logger.log("Bloques despues: " + blocksToRemove);
        if (blocksToRemove === 0) break;
        if ((currentBlocks - 1) <= minBlocks){
          Logger.log(candidates);
          Logger.log("Indice: " + i);
          Logger.log("destruimos: ");
          Logger.log(candidates[i]);
          candidates.splice(i, 1);
          Logger.log("post delete:");
          Logger.log(candidates);
          Logger.log("longitud candidatos: " + candidates.length);
          break;
        }   
      }
    }


    if (blocksToRemove > 0) {

      Logger.log("EventDayBlocks antes del tiempo extra");
      Logger.log(JSON.stringify(eventDayBlocks, null, 2));

      // Mapa auxiliar para poder ir de evId → meta (deadline, allowedDays, etc.)
      const metaById = {};
      eventsMeta.forEach(m => { metaById[m.id] = m; });

      // Distribuye numBlocks de tiempo extra de un evento concreto a lo largo de los días
      // permitidos hasta su deadline, intentando repartir en paquetes de 2 bloques por día
      // y balanceando para que no haya un día con mucha sobrecarga si otros aún tienen 0.
      function distributeExtraBlocks_(evId, originDayKey, numBlocks) {
        const meta = metaById[evId];
        // Si por lo que sea no tenemos meta, caemos al comportamiento antiguo (todo al día origen).
        if (!meta) {
          extraTimes[evId] = extraTimes[evId] || {
            concepto: eventDayBlocks[evId]["concepto"]
          };
          extraTimes[evId][originDayKey] =
            (extraTimes[evId][originDayKey] || 0) + numBlocks;
          return;
        }

        const deadlineKey = meta.deadline ? dateKeyFromDate_(meta.deadline) : null;
        const allowedRaw = (meta.allowedDays && meta.allowedDays.length)
          ? meta.allowedDays
          : dayKeys; // fallback: todas las dayKeys conocidas

        // Días candidatos: desde el día de origen hasta el deadline (incluido)
        const candidates = allowedRaw.filter(d =>
          d >= originDayKey && (!deadlineKey || d <= deadlineKey)
        );

        extraTimes[evId] = extraTimes[evId] || {
          concepto: eventDayBlocks[evId]["concepto"]
        };

        // Si no hay candidatos válidos, metemos todo en el propio día origen
        if (candidates.length === 0) {
          extraTimes[evId][originDayKey] =
            (extraTimes[evId][originDayKey] || 0) + numBlocks;
          return;
        }

        // Repartimos: mientras queden bloques, intentamos ponerlos en chunks de 2
        // y siempre en el día con menor carga extra actual.
        while (numBlocks > 0) {
          const chunk = (numBlocks >= 2) ? 2 : 1; // idealmente 2; si queda 1, lo añadimos encima

          // buscamos el día con menor carga extra actual para este evento
          let bestDay = candidates[0];
          let bestLoad = extraTimes[evId][bestDay] || 0;

          for (let i = 1; i < candidates.length; i++) {
            const d = candidates[i];
            const load = extraTimes[evId][d] || 0;
            if (load < bestLoad) {
              bestLoad = load;
              bestDay = d;
            }
          }

          extraTimes[evId][bestDay] =
            (extraTimes[evId][bestDay] || 0) + chunk;
          numBlocks -= chunk;
        }
      }

      // vamos sacando tiempo extra de ese día hasta 3h (6 bloques) o hasta agotar blocksToRemove
      let movedExtraBlocks = 0; // máximo 6 bloques (3h)

      while (blocksToRemove > 0 && movedExtraBlocks < 6) {
        // reconstruimos candidatos restantes del día
        const remaining = [];
        for (const evId in eventDayBlocks) {
          const blocks = eventDayBlocks[evId][dayKey] || 0;
          if (blocks > 0) {
            remaining.push({ evId, blocks });
          }
        }

        if (remaining.length === 0) break;

        // cogemos el evento con más bloques en ese día
        remaining.sort((a, b) => b.blocks - a.blocks);
        const { evId } = remaining[0];

        const available = eventDayBlocks[evId][dayKey] || 0;
        // no sacar más de lo disponible ni más de lo que falta por quitar ni más de 3h en total
        let toMove = Math.min(blocksToRemove, 6 - movedExtraBlocks, available);
        if (toMove <= 0) break;

        // opcional: evitar mover un solo bloque suelto si quieres mantener pares de 2
        // if (toMove === 1 && blocksToRemove > 1) toMove = 0;
        // if (toMove <= 0) break;

        // en vez de asignar todo al mismo dayKey, lo repartimos por días hasta el deadline
        distributeExtraBlocks_(evId, dayKey, toMove);

        eventDayBlocks[evId][dayKey] -= toMove;
        if (eventDayBlocks[evId][dayKey] <= 0) delete eventDayBlocks[evId][dayKey];

        dayUsedBlocks[dayKey] -= toMove;
        blocksToRemove -= toMove;
        movedExtraBlocks += toMove;
      }

      Logger.log("EventDayBlocks despues del tiempo extra");
      Logger.log(JSON.stringify(eventDayBlocks, null, 2));

      Logger.log("extraTimes:");
      Logger.log(JSON.stringify(extraTimes, null, 2));
    }


  });


  // Recalcular stats tras meter el extra
  stats = computeOverflowStats_(dayUsedBlocks);
  // Resultados finales
  const result = {
    freeMinutesByDay: freeMinutesByDay,
    finalDayUsedMinutes: Object.fromEntries(
      dayKeys.map(k => [k, (dayUsedBlocks[k] || 0) * 30])
    ),
    overflowFinal: stats.overflow,
    meanOverflowFinal: stats.meanOverflow,
    FFinal: stats.F,
    eventDayBlocks: eventDayBlocks,
    extraTimes: extraTimes,
    eventsMeta: eventsMeta,
    dayKeys: dayKeys
  };

  Logger.log("Resultado final optimización por torres (precalendario lógico):");
  Logger.log(JSON.stringify(result, null, 2));

  return result;
  
  // HELPER para sacar los overFlowStats
  // se tiene que meter aqui porque usa variables de dentro de la funcion y para no tener que meterlo en cada llamada
  function computeOverflowStats_(dayUsedBlocksSnapshot) {
    const dayMinutes = {};
    const overflow = {};
    let sumOverflow = 0;
    let sumAbsOverflow = 0;   
    let countDays = 0;

    dayKeys.forEach(dayKey => {
      const usedBlocks = dayUsedBlocksSnapshot[dayKey] || 0;
      const minutesAssigned = usedBlocks * 30;
      dayMinutes[dayKey] = minutesAssigned;
      const baseFree = freeMinutesByDay[dayKey] || 0;
      const ov = minutesAssigned - baseFree;
      overflow[dayKey] = ov;
      sumOverflow += ov;
      sumAbsOverflow += Math.abs(ov);  
      countDays++;
    });

    const meanOverflow = countDays > 0 ? (sumOverflow / countDays) : 0;

    let F = 0;
    dayKeys.forEach(dayKey => {
      const ov = overflow[dayKey];
      const diff = ov - meanOverflow;
      F += diff * diff;
    });

    return {
      dayMinutes: dayMinutes,
      overflow: overflow,
      meanOverflow: meanOverflow,
      F: F,
      sumAbsOverflow: sumAbsOverflow
    };
  }

  // helper para ver si ese dia es usable para ese día
  function isDayUsableForEvent_(meta, dayKey) {
  // Si no hay deadline (o no es una fecha válida), no restringimos
  if (!meta.deadline) return true;
  const deadline = new Date(meta.deadline);
  if (isNaN(deadline.getTime())) return true;

  // Solo nos importa el mismo día del deadline; otros días se tratan por fecha en allowedDays
  const deadlineKey = dateKeyFromDate_(deadline);
  if (deadlineKey !== dayKey) return true;

  const deadlineTime = deadline.getTime();

  // Buscamos si existe al menos un freeSlot en ese día que empiece antes del deadline
  for (let i = 0; i < freeSlots.length; i++) {
    const slot = freeSlots[i];
    const s = new Date(slot.start);
    const e = new Date(slot.end);
    if (dateKeyFromDate_(s) !== dayKey) continue;
    // Cualquier hueco que empiece antes del deadline se considera usable;
    // el ajuste fino de minutos ya lo hará scheduleTasksIntoFreeSlots_
    if (s.getTime() < deadlineTime && e.getTime() > s.getTime()) {
      return true;
    }
  }
  // No hay huecos reales antes del deadline en ese día → no usarlo en el precalendario
  return false;
}

}


// -------------- FUNCIONES SIN REVISAR

// NUEVO: horizonte dinámico basado en el último deadline relevante (Examen, Tareas, Laboratorio)
function getDynamicHorizonEnd_(now) {
  const candidates = [];

  // --- 1) Último deadline de TAREAS (hoja) ---
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(TASKS_SHEET_NAME);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= FIRST_TASK_ROW) {
        const numRows = lastRow - FIRST_TASK_ROW + 1;
        // Leemos concepto, deadline, eta y progreso (lo mínimo para filtrar tareas reales)
        const values = sheet.getRange(FIRST_TASK_ROW, 1, numRows, COL_PROGRESO).getValues();
        for (let i = 0; i < values.length; i++) {
          const row = values[i];
          const concepto = row[COL_CONCEPTO - 1];
          let deadline = row[COL_DEADLINE - 1];
          const etaHours = Number(row[COL_ETA - 1]) || 0;
          const progreso = (Number(row[COL_PROGRESO - 1]) || 0) / 100;

          if (!concepto || etaHours <= 0) continue;
          if (!(deadline instanceof Date) || isNaN(deadline.getTime())) continue;

          // Misma regla que en getPendingTasks_: deadline efectiva = día anterior 23:59
          const eff = new Date(deadline.getTime());
          eff.setDate(eff.getDate() - 1);
          eff.setHours(23, 59, 0, 0);

          const remainingMinutes = etaHours * 60 * (1 - progreso);
          if (remainingMinutes <= 0) continue;

          candidates.push(eff);
        }
      }
    }
  } catch (e) {
    Logger.log("getDynamicHorizonEnd_: fallo leyendo tareas: " + e);
  }

  // --- 2) Último EXAMEN (calendario) ---
  try {
    const calEx = fetchCalendar_(EXAM_CALENDAR_NAME);
    const start = startOfDay_(now);
    const end = addDays_(start, 365); // ventana amplia (puedes subir/bajar)
    const evs = calEx.getEvents(start, end);
    evs.forEach(ev => {
      const s = ev.getStartTime();
      if (s && s instanceof Date && !isNaN(s.getTime())) candidates.push(s);
    });
  } catch (e) {
    Logger.log("getDynamicHorizonEnd_: fallo leyendo examenes: " + e);
  }

  // --- 3) Último LABORATORIO (calendarios fijos) ---
  // Por tu config, puede existir "Laboratorio" y/o "Laboratorios"
  ["Laboratorio", "Laboratorios"].forEach(name => {
    try {
      const calLab = fetchCalendar_(name);
      const start = startOfDay_(now);
      const end = addDays_(start, 365);
      const evs = calLab.getEvents(start, end);
      evs.forEach(ev => {
        const e = ev.getEndTime();
        if (e && e instanceof Date && !isNaN(e.getTime())) candidates.push(e);
      });
    } catch (e) {
      // no hacemos nada: si no existe o falla, se ignora
    }
  });

  // Si no hay nada, fallback a comportamiento actual
  if (candidates.length === 0) {
    return new Date(now.getTime() + HORIZON_DAYS * 24 * 60 * 60 * 1000);
  }

  // Cogemos el más tardío y le damos un pequeño margen para que el planning lo “incluya bien”
  let maxD = candidates[0];
  for (let i = 1; i < candidates.length; i++) {
    if (candidates[i].getTime() > maxD.getTime()) maxD = candidates[i];
  }

  // margen: +1 día para que el último día tenga slots completos
  return addDays_(startOfDay_(maxD), 1);
}
