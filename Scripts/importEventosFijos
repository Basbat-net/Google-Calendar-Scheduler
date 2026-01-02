/************ CONFIG ************/

const ROW_START = 1497;
const ROW_END   = 2000;
const ROW_HEIGHT_PX = 21;

const IMPORT_WINDOW_DAYS = 365;   // window for Examenes scan
const FALLBACK_WEEKS_NO_EXAMS = 3;

// Where logs are WRITTEN now (no indexados)
const CALENDAR_WRITE_COL = {
  "Clases":        "I",
  "General":       "J",
  "Examenes":      "K",
  "Laboratorios":  "L"
};

// Where we CHECK / KEEP indexados
const CALENDAR_CHECK_COL = {
  "Clases":        "A",
  "General":       "B",
  "Examenes":      "C",
  "Laboratorios":  "D"
};

// --- NEW: Tareas columns ---
const TASKS_DB_WRITE_COL_LETTER = "M"; // no-indexados
const TASKS_DB_CHECK_COL_LETTER = "G"; // indexados (as you requested)

/**
 * Logs future events into no-indexados columns, deduping against indexados columns by FULL TEXT.
 *
 * NEW:
 * - If an event exists in indexados (A/B/C/D) but its start/end differs from the current calendar snapshot,
 *   remove it from indexados and append the updated version into no indexados (I/J/K/L).
 *
 * NEW 2 (your request):
 * - Apply the same "same id, changed times" reconciliation ALSO to the no-indexados (write) column:
 *   if an event is already present in I/J/K/L with the same id but different start/end from snapshot,
 *   remove the old version from I/J/K/L as well, and (re)append the updated version once.
 *
 * NEW 3 (your request):
 * - ALSO process "Tareas" into column M (same ROW_START..ROW_END), reconciling by Concepto:
 *   if same Concepto exists but any of Deadline/Prio/Prep/ETA/Progreso% differs, move corrected into M and delete old.
 *   Apply reconciliation both from G->M and within M itself.
 */
function logUpcomingCalendarEventsToDatabase() {
  const ss = SpreadsheetApp.getActive();
  const db = ss.getSheetByName(DATABASE_SHEET_NAME);
  if (!db) throw new Error("No existe la hoja '" + DATABASE_SHEET_NAME + "'.");

  const numRows = ROW_END - ROW_START + 1;
  db.setRowHeights(ROW_START, numRows, ROW_HEIGHT_PX);

  const now = new Date();

  // 1) compute cutoff windows based on Examenes
  const examInfo = getExamCutoff_(now);
  const examWindowEnd  = examInfo.examWindowEnd;
  const otherWindowEnd = examInfo.otherCalendarsEnd;

  // 2) process calendars
  Object.keys(CALENDAR_WRITE_COL).forEach(calendarName => {
    const writeColLetter = CALENDAR_WRITE_COL[calendarName];
    const checkColLetter = CALENDAR_CHECK_COL[calendarName];

    const writeColIndex = colFromA1_(writeColLetter);
    const checkColIndex = colFromA1_(checkColLetter);

    const cal = CalendarApp.getCalendarsByName(calendarName)[0];
    if (!cal) {
      Logger.log("Calendario no encontrado: " + calendarName);
      return;
    }

    const end = (calendarName === "Examenes") ? examWindowEnd : otherWindowEnd;
    const events = cal.getEvents(now, end);

    // Snapshot
    const currentById = new Map();
    const currentLogsSet = new Set();

    for (let i = 0; i < events.length; i++) {
      const ev = events[i];
      const s = ev.getStartTime();
      const e = ev.getEndTime();
      if (!(s instanceof Date) || isNaN(s)) continue;
      if (!(e instanceof Date) || isNaN(e)) continue;
      if (s.getTime() < now.getTime()) continue;

      const logText = String(calendarAppEventToPlainText_NoIndexados_(ev));
      const key = normalizeLogKey_(logText);
      const meta = parseLogMeta_(logText);

      currentLogsSet.add(key);

      if (meta.id) {
        currentById.set(meta.id, {
          logText,
          startMs: meta.startMs,
          endMs: meta.endMs
        });
      }
    }

    // --- READ indexados ---
    const checkRange = db.getRange(ROW_START, checkColIndex, numRows, 1);
    const checkVals = checkRange.getValues();

    const keptIndexados = [];
    const movedToNoIndexados = [];

    for (let i = 0; i < checkVals.length; i++) {
      const cell = checkVals[i][0];
      if (!cell) continue;

      const text = String(cell);
      const key = normalizeLogKey_(text);

      if (currentLogsSet.has(key)) {
        keptIndexados.push([text]);
        continue;
      }

      const oldMeta = parseLogMeta_(text);
      if (oldMeta.id && currentById.has(oldMeta.id)) {
        const cur = currentById.get(oldMeta.id);
        const changed =
          (isFiniteNumber_(oldMeta.startMs) && oldMeta.startMs !== cur.startMs) ||
          (isFiniteNumber_(oldMeta.endMs)   && oldMeta.endMs   !== cur.endMs);

        if (changed) {
          movedToNoIndexados.push([cur.logText]);
          continue;
        }

        keptIndexados.push([text]);
      }
    }

    writeCompactedColumn_(checkRange, keptIndexados);

    // --- READ write column ---
    const writeRange = db.getRange(ROW_START, writeColIndex, numRows, 1);
    writeRange.setWrap(false);
    const writeVals = writeRange.getValues();

    const keptWrite = [];
    const removedWriteBecauseChanged = [];

    for (let i = 0; i < writeVals.length; i++) {
      const cell = writeVals[i][0];
      if (!cell) continue;

      const text = String(cell);
      const key = normalizeLogKey_(text);

      if (currentLogsSet.has(key)) {
        keptWrite.push([text]);
        continue;
      }

      const oldMeta = parseLogMeta_(text);
      if (oldMeta.id && currentById.has(oldMeta.id)) {
        const cur = currentById.get(oldMeta.id);
        const changed =
          (isFiniteNumber_(oldMeta.startMs) && oldMeta.startMs !== cur.startMs) ||
          (isFiniteNumber_(oldMeta.endMs)   && oldMeta.endMs   !== cur.endMs);

        if (changed) {
          removedWriteBecauseChanged.push([cur.logText]);
          continue;
        }

        keptWrite.push([text]);
      }
    }

    writeCompactedColumn_(writeRange, keptWrite);

    const writeValsAfter = writeRange.getValues();
    let writeOffset = 0;
    while (writeOffset < writeValsAfter.length && writeValsAfter[writeOffset][0]) writeOffset++;
    if (writeOffset >= writeValsAfter.length) return;
    // esto corre solo si hay al menos uno
    // --- dedupe set ---
    const alreadyLogged = new Set();
    keptIndexados.forEach(r => alreadyLogged.add(normalizeLogKey_(r[0])));
    keptWrite.forEach(r => alreadyLogged.add(normalizeLogKey_(r[0])));

    const out = [];

    function pushIfNotLogged_(logText) {
      const key = normalizeLogKey_(logText);
      if (alreadyLogged.has(key)) return;
      if (writeOffset + out.length >= writeValsAfter.length) return;
      out.push([logText]);
      alreadyLogged.add(key);
    }

    movedToNoIndexados.forEach(r => pushIfNotLogged_(r[0]));
    removedWriteBecauseChanged.forEach(r => pushIfNotLogged_(r[0]));

    for (let i = 0; i < events.length; i++) {
      const ev = events[i];
      const s = ev.getStartTime();
      const e = ev.getEndTime();
      if (!(s instanceof Date) || isNaN(s)) continue;
      if (!(e instanceof Date) || isNaN(e)) continue;
      if (s.getTime() < now.getTime()) continue;

      const logText = String(calendarAppEventToPlainText_NoIndexados_(ev));
      const key = normalizeLogKey_(logText);

      if (alreadyLogged.has(key)) continue;
      if (writeOffset + out.length >= writeValsAfter.length) break;

      out.push([logText]);
      alreadyLogged.add(key);
    }

    if (out.length) {
      db.getRange(ROW_START + writeOffset, writeColIndex, out.length, 1).setValues(out);
    }
  });

  processTasksToDatabase_(db, now);
}

/**
 * Tareas import/reconcile:
 * - Snapshot tasks from TASKS_SHEET_NAME (pending only, like getPendingTasks_ base logic)
 * - Dedupe by FULL TEXT
 * - Reconcile by Concepto:
 *   If same Concepto but different Deadline/Prio/Prep/ETA/Progreso%, remove old and append corrected into M.
 * - Cleanup entries in G/M not present in snapshot.
 *
 * NEW (your request):
 * - When detecting "same Concepto but changed params" in INDEXADOS (G),
 *   only delete-from-indexados + move-to-no-indexados if programTime is from YESTERDAY
 *   relative to the run day (GMT+1 / script timezone).
 */
/**
 * Reemplaza (o añade) la línea "programTime: ..." en un texto de tarea.
 * - Si existe, la sustituye por programTime: <now>
 * - Si no existe, la añade al final.
 */
function replaceOrSetProgramTimeLine_(taskText, now) {
  const lines = String(taskText || "").split("\n");
  const newLine = "programTime: " + safeStr_(new Date(now.getTime()));

  let replaced = false;

  for (let i = 0; i < lines.length; i++) {
    if (lines[i].startsWith("programTime: ")) {
      lines[i] = newLine;
      replaced = true;
      break;
    }
  }

  if (!replaced) {
    lines.push(newLine);
  }

  return lines.join("\n");
}

function processTasksToDatabase_(db, now) {
  const numRows = ROW_END - ROW_START + 1;

  const writeColIndex = colFromA1_(TASKS_DB_WRITE_COL_LETTER); // M
  const checkColIndex = colFromA1_(TASKS_DB_CHECK_COL_LETTER); // G

  const snapshot = buildTasksSnapshot_(now);
  const currentLogsSet = snapshot.currentLogsSet;
  const currentByConcepto = snapshot.currentByConcepto;
  const orderedLogs = snapshot.orderedLogs;

  // ✅ dayKey “hoy” (zona local del script; GMT+1 si tu TZ es Europe/Madrid)
  const todayKey = dateKeyFromDate_(now);
  Logger.log("processTasksToDatabase_: now=%s todayKey=%s", now, todayKey);

  const checkRange = db.getRange(ROW_START, checkColIndex, numRows, 1);
  const checkVals = checkRange.getValues();

  const keptIndexados = [];
  const movedToNoIndexados = [];

  for (let i = 0; i < checkVals.length; i++) {
    const cell = checkVals[i][0];
    if (!cell) continue;

    const text = String(cell);

    const oldMeta = parseTaskMeta_(text);

    // ✅ LOGS: programTime evaluation
    Logger.log(
      "G[%s] concepto='%s' programTimeMs=%s programDayKey='%s' todayKey='%s'",
      (ROW_START + i),
      (oldMeta.concepto || ""),
      (isFiniteNumber_(oldMeta.programTimeMs) ? String(oldMeta.programTimeMs) : "NaN"),
      (oldMeta.programDayKey || ""),
      todayKey
    );

    // ✅ NEW TOP RULE (your request):
    // Si el Concepto existe en el snapshot actual (o sea, coincide con una tarea recuperada),
    // y programTime es de un día distinto a hoy, entonces se BORRA de G (se compacta hacia arriba).
    // No altera el resto de reglas: simplemente hace "continue" antes de la lógica existente.
    const hasProgramDay = !!oldMeta.programDayKey;
    const conceptoIsInSnapshot = !!oldMeta.concepto && currentByConcepto.has(oldMeta.concepto);
    const programDayIsNotToday = hasProgramDay && (oldMeta.programDayKey !== todayKey);

    if (conceptoIsInSnapshot && programDayIsNotToday) {
      Logger.log(
        "G[%s] DELETE (top rule): concepto='%s' is in snapshot and programDayKey='%s' != todayKey='%s'",
        (ROW_START + i),
        oldMeta.concepto,
        oldMeta.programDayKey,
        todayKey
      );
      continue; // NO se añade a keptIndexados => se elimina de G por compactado
    }

    const key = taskDedupeKeyFromText_(text);

    // Si coincide exactamente con el snapshot actual, se queda
    if (currentLogsSet.has(key)) {
      Logger.log("G[%s] KEEP (snapshot exact match by key)", (ROW_START + i));
      keptIndexados.push([text]);
      continue;
    }

    // ✅ FIX: si programTime NO es del mismo día que hoy, se borra de G.
    // Y si además existe en snapshot, empuja la versión actual a M.
    //
    // IMPORTANTE:
    // - Si NO podemos parsear programTime (programDayKey vacío), también se elimina de G
    //   (esto evita “atascos” cuando Date(v) falla y nunca se cumple la condición).
    const programDayUnknown = !hasProgramDay; // no parseable / missing

    Logger.log(
      "G[%s] programTime check: hasProgramDay=%s programDayUnknown=%s programDayIsNotToday=%s",
      (ROW_START + i),
      String(hasProgramDay),
      String(programDayUnknown),
      String(programDayIsNotToday)
    );

    if (programDayUnknown || programDayIsNotToday) {
      Logger.log("G[%s] DELETE from G due to programTime day mismatch/unknown", (ROW_START + i));

      if (oldMeta.concepto && currentByConcepto.has(oldMeta.concepto)) {
        const cur = currentByConcepto.get(oldMeta.concepto);
        Logger.log("G[%s] -> MOVE to M (snapshot current) concepto='%s'", (ROW_START + i), oldMeta.concepto);
        movedToNoIndexados.push([cur.logText]);
      } else {
        Logger.log("G[%s] -> NO MOVE to M (concepto not found in snapshot) concepto='%s'", (ROW_START + i), (oldMeta.concepto || ""));
      }

      continue; // NO se añade a keptIndexados => se elimina de G
    }

    // Si tenemos el mismo concepto en snapshot, aplicamos reconcile por cambios
    if (oldMeta.concepto && currentByConcepto.has(oldMeta.concepto)) {
      const cur = currentByConcepto.get(oldMeta.concepto);

      const changed = taskMetaChanged_(oldMeta, cur.meta);
      Logger.log(
        "G[%s] reconcile concepto='%s': changed=%s",
        (ROW_START + i),
        oldMeta.concepto,
        String(changed)
      );

      // Si mismo Concepto pero cambian campos “reales”, se borra de G y se empuja la versión actualizada a M
      if (changed) {
        Logger.log("G[%s] DELETE from G due to meta change; MOVE updated to M concepto='%s'", (ROW_START + i), oldMeta.concepto);
        movedToNoIndexados.push([cur.logText]);
        continue; // se elimina de indexados
      }

      // Si no cambió y es de HOY, se mantiene
      Logger.log("G[%s] KEEP (same-day programTime + no meta change) concepto='%s'", (ROW_START + i), oldMeta.concepto);
      keptIndexados.push([text]);
      continue;
    }

    // Si no está en snapshot (ya no existe), se elimina (no se conserva)
    Logger.log("G[%s] DELETE from G (not in snapshot by key and concepto not found)", (ROW_START + i));
  }

  // Esto SIEMPRE se ejecuta antes de cualquier return temprano posterior
  Logger.log("processTasksToDatabase_: compacting G; kept=%s movedToM=%s", keptIndexados.length, movedToNoIndexados.length);
  writeCompactedColumn_(checkRange, keptIndexados);

  const writeRange = db.getRange(ROW_START, writeColIndex, numRows, 1);
  writeRange.setWrap(false);
  let writeVals = writeRange.getValues();

  // ✅ NEW: Pre-pase en M (your request)
  // Si el concepto existe en el snapshot y programTime es de OTRO DÍA (horas irrelevantes),
  // entonces NO borra ni compacta; actualiza "programTime:" en la MISMA celda a "now".
  // (Se hace antes de dedupe/cleanup para que el resto de lógica opere con programTime actualizado.)
  let touchedM = false;

  for (let i = 0; i < writeVals.length; i++) {
    const cell = writeVals[i][0];
    if (!cell) continue;

    const text = String(cell);
    const meta = parseTaskMeta_(text);

    const hasProgramDay_M = !!meta.programDayKey;
    const conceptoIsInSnapshot_M = !!meta.concepto && currentByConcepto.has(meta.concepto);
    const programDayIsNotToday_M = hasProgramDay_M && (meta.programDayKey !== todayKey);

    if (conceptoIsInSnapshot_M && programDayIsNotToday_M) {
      const updated = replaceOrSetProgramTimeLine_(text, now);
      writeVals[i][0] = updated;
      touchedM = true;

      Logger.log(
        "M[%s] UPDATE programTime (same concepto in snapshot, different day): concepto='%s' oldDay='%s' newDay='%s'",
        (ROW_START + i),
        meta.concepto,
        meta.programDayKey,
        todayKey
      );
    }
  }

  if (touchedM) {
    writeRange.setValues(writeVals);
    // Relee para que el resto de la función trabaje con el contenido ya persistido
    writeVals = writeRange.getValues();
  }

  const keptWrite = [];
  const removedWriteBecauseChanged = [];

  // ✅ EXTRA: remove duplicates already present in M by stable key
  const seenWrite = new Set();

  for (let i = 0; i < writeVals.length; i++) {
    const cell = writeVals[i][0];
    if (!cell) continue;

    const text = String(cell);
    const key = taskDedupeKeyFromText_(text);

    // elimina duplicados ya existentes en M
    if (seenWrite.has(key)) {
      Logger.log("M[%s] DROP duplicate by key", (ROW_START + i));
      continue;
    }
    seenWrite.add(key);

    if (currentLogsSet.has(key)) {
      keptWrite.push([text]);
      continue;
    }

    const oldMeta = parseTaskMeta_(text);
    if (oldMeta.concepto && currentByConcepto.has(oldMeta.concepto)) {
      const cur = currentByConcepto.get(oldMeta.concepto);
      if (taskMetaChanged_(oldMeta, cur.meta)) {
        Logger.log("M[%s] REMOVE outdated; concepto='%s' -> will re-add updated", (ROW_START + i), oldMeta.concepto);
        removedWriteBecauseChanged.push([cur.logText]);
        continue;
      }
      keptWrite.push([text]);
    }
    // Si no está en snapshot, se elimina (no se conserva)
  }

  writeCompactedColumn_(writeRange, keptWrite);

  const writeValsAfter = writeRange.getValues();
  let writeOffset = 0;
  while (writeOffset < writeValsAfter.length && writeValsAfter[writeOffset][0]) writeOffset++;
  if (writeOffset >= writeValsAfter.length) return;

  const alreadyLogged = new Set();
  keptIndexados.forEach(r => alreadyLogged.add(taskDedupeKeyFromText_(r[0])));
  keptWrite.forEach(r => alreadyLogged.add(taskDedupeKeyFromText_(r[0])));

  const out = [];

  function pushIfNotLogged_(logText) {
    const key = taskDedupeKeyFromText_(logText);
    if (alreadyLogged.has(key)) return;
    if (writeOffset + out.length >= writeValsAfter.length) return;
    out.push([logText]);
    alreadyLogged.add(key);
  }

  movedToNoIndexados.forEach(r => pushIfNotLogged_(r[0]));
  removedWriteBecauseChanged.forEach(r => pushIfNotLogged_(r[0]));

  for (let i = 0; i < orderedLogs.length; i++) {
    const logText = orderedLogs[i];
    const key = taskDedupeKeyFromText_(logText);

    if (alreadyLogged.has(key)) continue;
    if (writeOffset + out.length >= writeValsAfter.length) break;

    out.push([logText]);
    alreadyLogged.add(key);
  }

  if (out.length) {
    importWeekFromToday();
    db.getRange(ROW_START + writeOffset, writeColIndex, out.length, 1).setValues(out);
  }
}






/**
 * Builds a snapshot of pending tasks using the same base logic as getPendingTasks_(now),
 * but producing:
 * - currentLogsSet: Set(fullLogText)
 * - currentByConcepto: Map(concepto -> { logText, meta })
 * - orderedLogs: logs in sheet order (for stable append)
 *
 * NEW:
 * - Each task log includes "programTime: <now>" so the DB entries can be gated later.
 */
function buildTasksSnapshot_(now) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TASKS_SHEET_NAME);
  if (!sheet) throw new Error("No existe la hoja de tareas '" + TASKS_SHEET_NAME + "'.");

  const lastRow = sheet.getLastRow();
  if (lastRow < FIRST_TASK_ROW) {
    return {
      currentLogsSet: new Set(),
      currentByConcepto: new Map(),
      orderedLogs: []
    };
  }

  const numRows = lastRow - FIRST_TASK_ROW + 1;
  const values = sheet.getRange(FIRST_TASK_ROW, COL_CONCEPTO, numRows, COL_PROGRESO).getValues();

  const currentLogsSet = new Set();
  const currentByConcepto = new Map();
  const orderedLogs = [];

  for (let i = 0; i < values.length; i++) {
    const rowIndex = FIRST_TASK_ROW + i;
    const row = values[i];

    const concepto = row[COL_CONCEPTO - 1];
    if (!concepto) continue;

    let progresoPct = Number(row[COL_PROGRESO - 1]);

    // ✅ NUEVO: ignorar tareas completadas (>= 100 %)
    if (!isFinite(progresoPct) || progresoPct >= 100) continue;

    let deadline = row[COL_DEADLINE - 1];
    if (deadline instanceof Date && !isNaN(deadline)) {
      const d = new Date(deadline.getTime());
      d.setDate(d.getDate() - 1);
      d.setHours(23, 59, 0, 0);
      deadline = d;
    }

    const prio = Number(row[COL_PRIO - 1]) || 0;
    const etaHours = Number(row[COL_ETA - 1]) || 0;

    if (!etaHours) continue;

    const logText = taskRowToPlainText_NoIndexados_({
      rowIndex,
      concepto,
      deadline,
      prio,
      prep: (typeof COL_PREP !== "undefined" && COL_PREP !== null)
        ? row[COL_PREP - 1]
        : "",
      etaHours,
      progresoPct,

      // ✅ NUEVO: timestamp de “programación”
      programTime: new Date(now.getTime())
    });

    const meta = parseTaskMeta_(logText);
    const key = taskDedupeKeyFromMeta_(meta);

    currentLogsSet.add(key);
    orderedLogs.push(logText);

    currentByConcepto.set(concepto, { logText, meta });
  }

  return { currentLogsSet, currentByConcepto, orderedLogs };
}

/**
 * Returns true if any of the user-specified fields differ:
 * Deadline, Prio, Prep, ETA, Progreso%
 *
 * IMPORTANT:
 * - programTime is NOT part of the diff check.
 */
function taskMetaChanged_(oldMeta, newMeta) {
  // Deadline
  if (isFiniteNumber_(oldMeta.deadlineMs) && isFiniteNumber_(newMeta.deadlineMs) && oldMeta.deadlineMs !== newMeta.deadlineMs) return true;

  // Prio
  if (isFiniteNumber_(oldMeta.prio) && isFiniteNumber_(newMeta.prio) && oldMeta.prio !== newMeta.prio) return true;

  // Prep (string compare)
  if (String(oldMeta.prep || "") !== String(newMeta.prep || "")) return true;

  // ETA (hours)
  if (isFiniteNumber_(oldMeta.etaHours) && isFiniteNumber_(newMeta.etaHours) && oldMeta.etaHours !== newMeta.etaHours) return true;

  // Progreso% (number)
  if (isFiniteNumber_(oldMeta.progresoPct) && isFiniteNumber_(newMeta.progresoPct) && oldMeta.progresoPct !== newMeta.progresoPct) return true;

  return false;
}

/**
 * Task plain text formatter.
 * IMPORTANT: Must NOT start with "=" or "===" (Sheets would treat as formula).
 */
function taskRowToPlainText_NoIndexados_(t) {
  const lines = [];
  lines.push("[TASK]");

  lines.push("concepto: " + safeStr_(t.concepto));
  lines.push("rowIndex: " + safeStr_(t.rowIndex));
  lines.push("deadline: " + safeStr_(t.deadline));
  lines.push("prio: " + safeStr_(t.prio));
  lines.push("prep: " + safeStr_(t.prep));
  lines.push("etaHours: " + safeStr_(t.etaHours));
  lines.push("progresoPct: " + safeStr_(t.progresoPct));

  // ✅ NUEVO: cuándo fue programada (para reglas de “solo mover si es de ayer”)
  lines.push("programTime: " + safeStr_(t.programTime));

  return lines.join("\n");
}

/**
 * Extracts concepto/deadline/prio/prep/eta/progreso/programTime from task log text.
 * We do NOT use this for dedupe; only to detect "same Concepto changed fields" + gating by programTime.
 *
 * Expects lines like:
 *   concepto: ...
 *   deadline: ...
 *   prio: ...
 *   prep: ...
 *   etaHours: ...
 *   progresoPct: ...
 *   programTime: ...
 */
function parseTaskMeta_(logText) {
  const res = {
    concepto: "",
    deadlineMs: NaN,
    prio: NaN,
    prep: "",
    etaHours: NaN,
    progresoPct: NaN,

    // ✅ NUEVO
    programTimeMs: NaN,
    programDayKey: "" // YYYY-MM-DD (zona local)
  };

  if (!logText) return res;
  const lines = String(logText).split("\n");

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    if (line.startsWith("concepto: ")) {
      res.concepto = line.substring(10).trim();
      continue;
    }
    if (line.startsWith("deadline: ")) {
      const v = line.substring(10).trim();
      const d = new Date(v);
      if (!isNaN(d.getTime())) res.deadlineMs = d.getTime();
      continue;
    }
    if (line.startsWith("prio: ")) {
      const v = line.substring(6).trim();
      const n = Number(v);
      if (isFinite(n)) res.prio = n;
      continue;
    }
    if (line.startsWith("prep: ")) {
      res.prep = line.substring(6).trim();
      continue;
    }
    if (line.startsWith("etaHours: ")) {
      const v = line.substring(9).trim();
      const n = Number(v);
      if (isFinite(n)) res.etaHours = n;
      continue;
    }
    if (line.startsWith("progresoPct: ")) {
      const v = line.substring(13).trim();
      const n = Number(v);
      if (isFinite(n)) res.progresoPct = n;
      continue;
    }

    // ✅ NUEVO (robusto):
    // Acepta:
    //  - "programTime: 1766791323000" (epoch ms)
    //  - "programTime: Sat Dec 26 2025 22:42:03 GMT+0100 (Central European Standard Time)"
    if (line.startsWith("programTime: ")) {
      const vRaw = line.substring(13).trim();

      // 1) intento numérico directo (epoch ms)
      const asNum = Number(vRaw);
      if (isFinite(asNum) && asNum > 0) {
        res.programTimeMs = asNum;
        const dNum = new Date(asNum);
        if (!isNaN(dNum.getTime())) res.programDayKey = dateKeyFromDate_(dNum);
        continue;
      }

      // 2) intento Date(vRaw) normal
      let d = new Date(vRaw);

      // 3) fallback: recorta lo de "(Central ...)" para que Date() lo trague mejor si falla
      if (isNaN(d.getTime())) {
        const vTrim = vRaw.replace(/\s*\(.*\)\s*$/, "").trim(); // quita "(...)"
        d = new Date(vTrim);
      }

      if (!isNaN(d.getTime())) {
        res.programTimeMs = d.getTime();
        res.programDayKey = dateKeyFromDate_(d);
      }
      continue;
    }
  }

  return res;
}


/**
 * Computes:
 * - examWindowEnd: now + IMPORT_WINDOW_DAYS
 * - otherCalendarsEnd: latest exam end (future exams), else now + 3 weeks
 */
function getExamCutoff_(now) {
  const cal = CalendarApp.getCalendarsByName("Examenes")[0];
  if (!cal) throw new Error("No se encontró el calendario 'Examenes'.");

  const examWindowEnd = new Date(now.getTime() + IMPORT_WINDOW_DAYS * 24 * 60 * 60 * 1000);
  const examEventsAll = cal.getEvents(now, examWindowEnd);

  let latestExamEnd = null;

  for (let i = 0; i < examEventsAll.length; i++) {
    const ev = examEventsAll[i];
    const s = ev.getStartTime();
    const e = ev.getEndTime();
    if (!(s instanceof Date) || isNaN(s.getTime())) continue;
    if (!(e instanceof Date) || isNaN(e.getTime())) continue;
    if (s.getTime() < now.getTime()) continue;

    if (!latestExamEnd || e.getTime() > latestExamEnd.getTime()) {
      latestExamEnd = new Date(e.getTime());
    }
  }

  let otherCalendarsEnd;
  if (latestExamEnd) {
    otherCalendarsEnd = latestExamEnd;
  } else {
    otherCalendarsEnd = new Date(now.getTime() + FALLBACK_WEEKS_NO_EXAMS * 7 * 24 * 60 * 60 * 1000);
  }

  return {
    examWindowEnd: examWindowEnd,
    otherCalendarsEnd: otherCalendarsEnd
  };
}

/**
 * Rewrites a 1-col range with kept rows at top, blanks below.
 */
function writeCompactedColumn_(range, keptRows) {
  const numRows = range.getNumRows();
  const out = [];

  for (let i = 0; i < numRows; i++) {
    out.push(i < keptRows.length ? keptRows[i] : [""]);
  }

  range.setValues(out);
}

/**
 * Extracts id/start/end from your full log text.
 * We do NOT use this for dedupe; only to detect "same event changed times".
 *
 * Expects lines like:
 *   id: ...
 *   start: ...
 *   end: ...
 */
function parseLogMeta_(logText) {
  const res = { id: "", startMs: NaN, endMs: NaN };

  if (!logText) return res;
  const lines = String(logText).split("\n");

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    if (line.startsWith("id: ")) {
      res.id = line.substring(4).trim();
      continue;
    }
    if (line.startsWith("start: ")) {
      const v = line.substring(7).trim();
      const d = new Date(v);
      if (!isNaN(d.getTime())) res.startMs = d.getTime();
      continue;
    }
    if (line.startsWith("end: ")) {
      const v = line.substring(5).trim();
      const d = new Date(v);
      if (!isNaN(d.getTime())) res.endMs = d.getTime();
      continue;
    }
  }

  return res;
}

function isFiniteNumber_(x) {
  return typeof x === "number" && isFinite(x);
}

/**
 * Plain text formatter.
 * IMPORTANT: Must NOT start with "=" or "===" (Sheets would treat as formula).
 */
function calendarAppEventToPlainText_NoIndexados_(ev) {
  const lines = [];
  lines.push("[GOOGLE CALENDAR EVENT]");

  lines.push("title: " + safeStr_(ev.getTitle()));
  lines.push("id: " + safeStr_(ev.getId()));
  lines.push("start: " + safeStr_(ev.getStartTime()));
  lines.push("end: " + safeStr_(ev.getEndTime()));
  lines.push("allDay: " + safeStr_(ev.isAllDayEvent()));
  lines.push("location: " + safeStr_(ev.getLocation()));
  lines.push("created: " + safeStr_(ev.getDateCreated()));
  lines.push("lastUpdated: " + safeStr_(ev.getLastUpdated()));
  lines.push("description:");
  lines.push(safeStr_(ev.getDescription()));

  return lines.join("\n");
}

function safeStr_(v) {
  if (v === null || v === undefined) return "";
  return String(v);
}

function colFromA1_(letter) {
  const s = String(letter).toUpperCase().trim();
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n;
}

function normalizeLogKey_(t) {
  return String(t || "")
    .replace(/\r\n/g, "\n")        // normalize line endings
    .replace(/[ \t]+$/gm, "")     // trim trailing spaces per line
    .trimEnd();                   // remove trailing newlines/spaces at end
}

/************ NEW: TASK DEDUPE KEYS ************/

/**
 * Key estable para dedupe de tareas (IGNORA programTime).
 * Usa el meta parseado: concepto + deadlineMs + prio + prep + etaHours + progresoPct
 * Si no hay concepto, fallback a FULL TEXT normalizado.
 */
function taskDedupeKeyFromText_(logText) {
  const m = parseTaskMeta_(logText);

  if (m.concepto) {
    const dl = isFiniteNumber_(m.deadlineMs) ? m.deadlineMs : "NaN";
    const pr = isFiniteNumber_(m.prio) ? m.prio : "NaN";
    const et = isFiniteNumber_(m.etaHours) ? m.etaHours : "NaN";
    const pg = isFiniteNumber_(m.progresoPct) ? m.progresoPct : "NaN";
    const pp = String(m.prep || "");

    return "TSK|" + m.concepto + "|" + dl + "|" + pr + "|" + pp + "|" + et + "|" + pg;
  }

  return "TSKTXT|" + normalizeLogKey_(logText);
}

/**
 * Misma key pero a partir de la meta snapshot.
 */
function taskDedupeKeyFromMeta_(meta) {
  if (!meta || !meta.concepto) return "TSKTXT|";
  const dl = isFiniteNumber_(meta.deadlineMs) ? meta.deadlineMs : "NaN";
  const pr = isFiniteNumber_(meta.prio) ? meta.prio : "NaN";
  const et = isFiniteNumber_(meta.etaHours) ? meta.etaHours : "NaN";
  const pg = isFiniteNumber_(meta.progresoPct) ? meta.progresoPct : "NaN";
  const pp = String(meta.prep || "");
  return "TSK|" + meta.concepto + "|" + dl + "|" + pr + "|" + pp + "|" + et + "|" + pg;
}
