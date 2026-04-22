/**
 * Loguea:
 *  1) Inputs diarios (Num / Bool / Hour) desde A:C (filas 4–49) al dataset "Inputs"
 *     con estructura: (timestamp, concepto, value)
 *     y resetea SOLO C4:C49 (bool->FALSE, num->"", hour->"") para filas con Concepto.
 *
 *  2) Seguimiento calendario desde D:H (filas 4–49) al dataset "Tareas y calendario"
 *     con estructura: (date, concepto, tipo, startTime, endTime, durationHours, doneBool)
 *     y SOLO destilda los checkboxes H4:H49 (si había concepto en D).
 *
 * DATABASE:
 *   Hoja: "database"
 *   Tabla "Ubicaciones sub datasets":
 *     Col A: Concepto (ej. "Inputs", "Tareas y calendario")
 *     Col B: Inicio
 *     Col C: Final
 *     Col D: First blank
 *
 * NUEVO (CADENCIAS):
 *   - Permite procesar [Med] y no-[Med] con flags opcionales:
 *     logDailyTrackingToDatabase("MainDashboard", { processMed: true, processNonMed: false })
 *   - Si una fila se ignora por flags, no se loguea y NO se resetea (actúa como si no existiera).
 */
function logDailyTrackingToDatabase(sheetName = "MainDashboard", opts) {
  opts = opts || {};
  const processMed = (opts.processMed !== false);          // default true
  const processNonMed = (opts.processNonMed !== false);    // default true

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  /************ CONFIG ORIGEN ************/
  const SOURCE_SHEET_NAME = "MainDashboard";

  const START_ROW = 4;
  const END_ROW = 49;

  // Inputs block: A:C
  const INPUT_COL_START = 1;     // A
  const INPUT_NUM_COLS = 3;      // A:C
  const INPUT_CONCEPTO_OFFSET = 0; // A
  const INPUT_TIPO_OFFSET = 1;     // B
  const INPUT_VALOR_OFFSET = 2;    // C

  // Seguimiento calendario block: D:H
  const CAL_COL_START = 4;       // D
  const CAL_NUM_COLS = 5;        // D:H
  const CAL_CONCEPTO_OFFSET = 0; // D
  const CAL_TIPO_OFFSET = 1;     // E
  const CAL_INICIO_OFFSET = 2;   // F
  const CAL_FINAL_OFFSET = 3;    // G
  const CAL_DONE_OFFSET = 4;     // H

  /************ CONFIG DATABASE ************/
  const DB_SHEET_NAME = "database";

  // Tabla "Ubicaciones sub datasets"
  const LOC_TABLE_START_ROW = 4;
  const LOC_CONCEPTO_COL = 1;      // A
  const LOC_INICIO_COL = 2;        // B
  const LOC_FINAL_COL = 3;         // C
  const LOC_FIRSTBLANK_COL = 4;    // D

  const DATASET_INPUTS = "Inputs";
  const DATASET_CAL = "Tareas y calendario";

  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) throw new Error(`No existe la hoja "${SOURCE_SHEET_NAME}"`);

  const dbSheet = ss.getSheetByName(DB_SHEET_NAME);
  if (!dbSheet) throw new Error(`No existe la hoja "${DB_SHEET_NAME}"`);

  const numRows = END_ROW - START_ROW + 1;

  /************ 1) INPUTS: LEER A:C (4–49) ************/
  const inputRange = sourceSheet.getRange(START_ROW, INPUT_COL_START, numRows, INPUT_NUM_COLS);
  const inputValues = inputRange.getValues();

  const inputRowsToLog = [];
  const inputRowsToReset = [];

  for (let i = 0; i < inputValues.length; i++) {
    const concepto = inputValues[i][INPUT_CONCEPTO_OFFSET];
    const tipo = inputValues[i][INPUT_TIPO_OFFSET];
    const valor = inputValues[i][INPUT_VALOR_OFFSET];

    if (concepto === "" || concepto === null) continue;

    const conceptoStr = String(concepto);

    // ✅ CAMBIO CLAVE: [Med] se ignora SIEMPRE en el "daily thing"
    // (se trata como fila inexistente: no log y no reset), aunque processMed sea true.
    if (conceptoStr.indexOf("[Med]") !== -1) continue;

    const isMed = false; // queda explícito para no romper el resto del flujo

    // NUEVO: cadencias por categoría
    // - Si es [Med] pero no toca procesar meds => ignorar totalmente (no log, no reset)
    if (isMed && !processMed) continue;
    // - Si NO es [Med] pero no toca procesar no-meds => ignorar totalmente (no log, no reset)
    if (!isMed && !processNonMed) continue;

    // Nuevo: soportar tipo "Hour" en col B
    // Normalizar tipo
    const tipoNorm = String(tipo || "").trim().toLowerCase();
    // ✅ NUEVO: si tipo es num u hour y C está vacío → ignorar completamente
    if ((tipoNorm === "num" || tipoNorm === "hour" || tipoNorm === "bool") && (valor === "" || valor === null || valor === false)) {
      continue;
    }

    // Por defecto, se loguea lo que haya en C
    let valueToLog = valor;

    // Conversión de Hour → horas decimales
    if (tipoNorm === "hour") {
      const frac = toTimeFraction_(valor); // Date | number | "HH:MM"
      if (frac === null) continue; // seguridad extra
      valueToLog = frac * 24;
    }

    inputRowsToLog.push([concepto, tipo, valueToLog]);
    inputRowsToReset.push({
      row: START_ROW + i,
      tipo: tipo,
      valor: valor
    });

  }

  /************ 2) CALENDARIO: LEER D:H (4–49) ************/
  const calRange = sourceSheet.getRange(START_ROW, CAL_COL_START, numRows, CAL_NUM_COLS);
  const calValues = calRange.getValues();

  const calRowsToLog = [];
  const calRowsToUntick = [];

  for (let i = 0; i < calValues.length; i++) {
    const concepto = calValues[i][CAL_CONCEPTO_OFFSET];
    const tipo = calValues[i][CAL_TIPO_OFFSET];
    const inicio = calValues[i][CAL_INICIO_OFFSET];
    const fin = calValues[i][CAL_FINAL_OFFSET];
    const done = calValues[i][CAL_DONE_OFFSET];

    if (concepto === "" || concepto === null) continue;

    // Nota: aquí NO filtramos por [Med] porque [Med] es un marcador de Inputs (col A),
    // y el bloque calendario es otro bloque distinto.
    // Si quisieras filtrar también en calendario, se puede añadir con el mismo patrón.

    calRowsToLog.push({
      row: START_ROW + i,
      concepto,
      tipo,
      inicio,
      fin,
      done: done === true
    });

    calRowsToUntick.push(START_ROW + i);
  }

  // Si no hay nada que loguear en ningún bloque, salimos
  if (inputRowsToLog.length === 0 && calRowsToLog.length === 0) return;

  /************ 3) WRITE INPUTS DATASET ************/
  if (inputRowsToLog.length > 0) {
    const locInputs = getDatasetLocation_(dbSheet, {
      datasetName: DATASET_INPUTS,
      locTableStartRow: LOC_TABLE_START_ROW,
      conceptoCol: LOC_CONCEPTO_COL,
      inicioCol: LOC_INICIO_COL,
      finalCol: LOC_FINAL_COL,
      firstBlankCol: LOC_FIRSTBLANK_COL
    });

    const timestamp = new Date();

    // NUEVO: si el concepto (col A) es EXACTAMENTE "Hora acostarse",
    // se guarda con offset temporal de -1 día.
    const rows = inputRowsToLog.map(r => {
      const concepto = r[0];
      const valueToLog = r[2];

      let ts = timestamp;
      if (concepto === "Hora acostarse") {
        ts = new Date(timestamp.getTime());
        ts.setDate(ts.getDate() - 1);
      }

      return [ts, concepto, valueToLog]; // (timestamp, concepto, value)
    });

    const lastNeededRow = locInputs.firstBlank + rows.length - 1;
    if (lastNeededRow > locInputs.fin) {
      throw new Error(`El dataset "${DATASET_INPUTS}" no tiene espacio suficiente`);
    }

    dbSheet.getRange(locInputs.firstBlank, 1, rows.length, 3).setValues(rows);
    dbSheet.getRange(locInputs.locRow, LOC_FIRSTBLANK_COL).setValue(locInputs.firstBlank + rows.length);

    // Reset SOLO C4:C49, pero solo en filas que realmente se procesaron (inputRowsToReset)
    // Nota: si una fila fue ignorada por flags (p. ej. no-[Med] fuera de la ventana diaria),
    // no entra en inputRowsToReset y NO se resetea.
    inputRowsToReset.forEach(item => {
      const tipo = String(item.tipo || "").trim().toLowerCase();
      let newVal;

      if (tipo === "bool") newVal = false;
      else if (tipo === "num") newVal = "";
      else if (tipo === "hour") newVal = "";
      else newVal = (typeof item.valor === "boolean") ? false : "";

      if (item.row >= START_ROW && item.row <= END_ROW) {
        sourceSheet.getRange(item.row, 3).setValue(newVal); // Col C
      }
    });
  }

  /************ 4) WRITE CALENDAR DATASET ("Tareas y calendario") ************/
  if (calRowsToLog.length > 0) {
    const locCal = getDatasetLocation_(dbSheet, {
      datasetName: DATASET_CAL,
      locTableStartRow: LOC_TABLE_START_ROW,
      conceptoCol: LOC_CONCEPTO_COL,
      inicioCol: LOC_INICIO_COL,
      finalCol: LOC_FINAL_COL,
      firstBlankCol: LOC_FIRSTBLANK_COL
    });

    const dateOnly = new Date();
    dateOnly.setHours(0, 0, 0, 0);

    const rows = calRowsToLog.map(r => {
      const durHours = calcDurationHours_(r.inicio, r.fin);
      return [
        dateOnly,         // date
        r.concepto,       // concepto
        r.tipo,           // tipo
        r.inicio,         // startTime
        r.fin,            // endTime
        durHours,         // duration hours
        r.done            // checkbox boolean
      ];
    });

    const lastNeededRow = locCal.firstBlank + rows.length - 1;
    if (lastNeededRow > locCal.fin) {
      throw new Error(`El dataset "${DATASET_CAL}" no tiene espacio suficiente`);
    }

    // Escribimos en A:G (7 columnas)
    dbSheet.getRange(locCal.firstBlank, 1, rows.length, 7).setValues(rows);
    dbSheet.getRange(locCal.locRow, LOC_FIRSTBLANK_COL).setValue(locCal.firstBlank + rows.length);

    // SOLO destildar checkboxes H4:H49 (solo filas con Concepto en D)
    // H es la columna 8
    calRowsToUntick.forEach(row => {
      if (row >= START_ROW && row <= END_ROW) {
        sourceSheet.getRange(row, 8).setValue(false); // Col H
      }
    });
  }
}

/**
 * Devuelve ubicación del dataset en la tabla "Ubicaciones sub datasets"
 * y los parámetros: inicio/fin/firstBlank/locRow.
 */
function getDatasetLocation_(dbSheet, cfg) {
  const locLastRow = dbSheet.getLastRow();
  const numRows = locLastRow - cfg.locTableStartRow + 1;
  if (numRows <= 0) throw new Error(`Tabla de ubicaciones vacía o mal configurada.`);

  const locRange = dbSheet.getRange(
    cfg.locTableStartRow,
    cfg.conceptoCol,
    numRows,
    cfg.firstBlankCol
  );
  const locValues = locRange.getValues();

  let idx = -1;
  for (let i = 0; i < locValues.length; i++) {
    if (String(locValues[i][0]).trim() === cfg.datasetName) {
      idx = i;
      break;
    }
  }
  if (idx === -1) throw new Error(`Dataset "${cfg.datasetName}" no encontrado en la tabla de ubicaciones.`);

  const locRow = cfg.locTableStartRow + idx;
  const inicio = Number(locValues[idx][cfg.inicioCol - 1]);
  const fin = Number(locValues[idx][cfg.finalCol - 1]);
  let firstBlank = Number(locValues[idx][cfg.firstBlankCol - 1]);

  if (!firstBlank || isNaN(firstBlank)) {
    firstBlank = findFirstBlankRowInRange_(dbSheet, inicio, fin, 1);
  }

  return { locRow, inicio, fin, firstBlank };
}

/**
 * Devuelve la primera fila vacía ("" o null) en un rango
 */
function findFirstBlankRowInRange_(sheet, startRow, endRow, col) {
  const values = sheet
    .getRange(startRow, col, endRow - startRow + 1, 1)
    .getValues();

  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === "" || values[i][0] === null) {
      return startRow + i;
    }
  }
  return endRow + 1;
}

/**
 * Calcula duración en horas entre dos celdas que pueden ser:
 * - Date (hora válida)
 * - número (fracción de día en Sheets)
 * - string tipo "13:00"
 *
 * Si falta inicio o fin, devuelve "" (vacío).
 */
function calcDurationHours_(startVal, endVal) {
  const s = toTimeFraction_(startVal);
  const e = toTimeFraction_(endVal);

  if (s === null || e === null) return "";

  let diff = e - s;
  if (diff < 0) diff += 1; // por si cruza medianoche

  return diff * 24;
}

/**
 * Convierte valores de hora a fracción de día (Sheets).
 * Retorna null si no se puede.
 */
function toTimeFraction_(v) {
  if (v === "" || v === null) return null;

  if (v instanceof Date) {
    return (v.getHours() * 3600 + v.getMinutes() * 60 + v.getSeconds()) / 86400;
  }

  if (typeof v === "number") {
    // Sheets time as fraction-of-day
    return v;
  }

  if (typeof v === "string") {
    const m = v.trim().match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (!m) return null;
    const hh = Number(m[1]);
    const mm = Number(m[2]);
    const ss = m[3] ? Number(m[3]) : 0;
    if (hh < 0 || hh > 23 || mm < 0 || mm > 59 || ss < 0 || ss > 59) return null;
    return (hh * 3600 + mm * 60 + ss) / 86400;
  }

  return null;
}

/************ PERSISTENCIA Y SCHEDULER (NUEVO) ************/

/**
 * Estado persistente: guarda/lee milisegundos en ScriptProperties.
 */
function getState_(key) {
  const props = PropertiesService.getScriptProperties();
  const v = props.getProperty(key);
  return v ? Number(v) : 0; // millis since epoch
}

function setState_(key, millis) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(key, String(millis));
}

function hoursSince_(millis) {
  if (!millis) return Infinity;
  return (Date.now() - millis) / (1000 * 60 * 60);
}

/**
 * Ejecuta:
 *  - [Med] cada 30 min (siempre)
 *  - No-[Med] 1 vez cada 24h (según estado persistente)
 */
function logTrackingScheduler_() {
  const LAST_DAILY_KEY = "LAST_DAILY_NONMED_LOG_MS";

  const lastDaily = getState_(LAST_DAILY_KEY);
  const shouldRunDailyNonMed = hoursSince_(lastDaily) >= 24;

  logDailyTrackingToDatabase("MainDashboard", {
    processMed: true,
    processNonMed: shouldRunDailyNonMed
  });

  if (shouldRunDailyNonMed) {
    setState_(LAST_DAILY_KEY, Date.now());
  }
}

function logMedTimeInputsToDatabase(sheetName = "MainDashboard") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  /************ CONFIG ORIGEN ************/
  const SOURCE_SHEET_NAME = "MainDashboard";

  const START_ROW = 4;
  const END_ROW = 49;

  // Inputs block: A:C
  const INPUT_COL_START = 1;        // A
  const INPUT_NUM_COLS = 3;         // A:C
  const INPUT_CONCEPTO_OFFSET = 0;  // A
  const INPUT_VALOR_OFFSET = 2;     // C

  /************ CONFIG DATABASE ************/
  const DB_SHEET_NAME = "database";

  // Tabla "Ubicaciones sub datasets"
  const LOC_TABLE_START_ROW = 4;
  const LOC_CONCEPTO_COL = 1;      // A
  const LOC_INICIO_COL = 2;        // B
  const LOC_FINAL_COL = 3;         // C
  const LOC_FIRSTBLANK_COL = 4;    // D

  const DATASET_INPUTS = "Inputs";

  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) throw new Error(`No existe la hoja "${SOURCE_SHEET_NAME}"`);

  const dbSheet = ss.getSheetByName(DB_SHEET_NAME);
  if (!dbSheet) throw new Error(`No existe la hoja "${DB_SHEET_NAME}"`);

  const numRows = END_ROW - START_ROW + 1;

  /************ LEER A:C (4–49) ************/
  const inputRange = sourceSheet.getRange(START_ROW, INPUT_COL_START, numRows, INPUT_NUM_COLS);
  const inputValues = inputRange.getValues();

  const rowsToLog = [];
  const rowsToReset = [];

  // Convierte hora a horas decimales (HH + MM/60 + SS/3600)
  function parseToDecimalHours_(v) {
    if (v === "" || v === null) return null;

    if (v instanceof Date) {
      return v.getHours() + (v.getMinutes() / 60) + (v.getSeconds() / 3600);
    }

    if (typeof v === "number" && !isNaN(v)) {
      const totalSeconds = Math.round(v * 86400);
      const hh = Math.floor(totalSeconds / 3600) % 24;
      const mm = Math.floor((totalSeconds % 3600) / 60);
      const ss = totalSeconds % 60;
      return hh + (mm / 60) + (ss / 3600);
    }

    if (typeof v === "string") {
      const m = v.trim().match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
      if (!m) return null;
      const hh = Number(m[1]);
      const mm = Number(m[2]);
      const ss = m[3] ? Number(m[3]) : 0;
      return hh + (mm / 60) + (ss / 3600);
    }

    return null;
  }

  for (let i = 0; i < inputValues.length; i++) {
    const concepto = inputValues[i][INPUT_CONCEPTO_OFFSET];
    const valor = inputValues[i][INPUT_VALOR_OFFSET];

    if (concepto === "" || concepto === null) continue;
    if (valor === "" || valor === null) continue; // ✅ SOLO si C tiene contenido

    const conceptoStr = String(concepto);
    if (conceptoStr.indexOf("[Med]") === -1) continue; // SOLO [Med]

    const decHours = parseToDecimalHours_(valor);
    if (decHours === null) continue;

    rowsToLog.push([conceptoStr, decHours]);
    rowsToReset.push(START_ROW + i); // marcar para limpiar C
  }

  if (rowsToLog.length === 0) return;

  /************ UBICACIÓN DATASET ************/
  const locInputs = getDatasetLocation_(dbSheet, {
    datasetName: DATASET_INPUTS,
    locTableStartRow: LOC_TABLE_START_ROW,
    conceptoCol: LOC_CONCEPTO_COL,
    inicioCol: LOC_INICIO_COL,
    finalCol: LOC_FINAL_COL,
    firstBlankCol: LOC_FIRSTBLANK_COL
  });

  const timestamp = new Date();

  const rows = rowsToLog.map(r => [timestamp, r[0], r[1]]); // (timestamp, concepto, value)

  const lastNeededRow = locInputs.firstBlank + rows.length - 1;
  if (lastNeededRow > locInputs.fin) {
    throw new Error(`El dataset "${DATASET_INPUTS}" no tiene espacio suficiente`);
  }

  dbSheet.getRange(locInputs.firstBlank, 1, rows.length, 3).setValues(rows);
  dbSheet.getRange(locInputs.locRow, LOC_FIRSTBLANK_COL).setValue(locInputs.firstBlank + rows.length);

  /************ LIMPIAR SOLO C EN FILAS [Med] PROCESADAS ************/
  rowsToReset.forEach(r => {
    if (r >= START_ROW && r <= END_ROW) {
      sourceSheet.getRange(r, 3).setValue(""); // Col C
    }
  });
}
