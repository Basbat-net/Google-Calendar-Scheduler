// BOOL epoch: days without data from this date are treated as FALSE (0)
const BOOL_EPOCH = new Date(2025, 11, 26); // months are 0-based -> 11 = December

function extractDatabaseDataAsDictAndLog_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("database");
  if (!sh) throw new Error("No existe la hoja 'database'.");

  // --- 1) Rango de filas ---
  const startRowBase = sh.getRange("B7").getValue();
  const endRowBase   = sh.getRange("D7").getValue();

  if (typeof startRowBase !== "number" || typeof endRowBase !== "number") {
    throw new Error("database!B7 y database!D7 deben ser numéricos.");
  }

  const startRow = Math.floor(startRowBase) + 1;
  const endRow   = Math.floor(endRowBase) - 1;

  if (endRow < startRow) {
    throw new Error(`Rango inválido: startRow=${startRow}, endRow=${endRow}`);
  }

  const numRows = endRow - startRow + 1;

  // --- 2) Detectar última columna ---
  const maxCols = sh.getLastColumn();
  const firstRowValues = sh.getRange(startRow, 1, 1, maxCols).getValues()[0];

  let lastCol = 0;
  for (let i = 0; i < firstRowValues.length; i++) {
    if (firstRowValues[i] === "" || firstRowValues[i] === null) break;
    lastCol = i + 1;
  }

  if (lastCol === 0) {
    throw new Error(`No hay datos en la fila ${startRow}`);
  }

  // --- 3) Extraer datos ---
  const rawData = sh.getRange(startRow, 1, numRows, lastCol).getValues();

  // --- 4) Construir diccionario (fecha truncada) ---
  const dict = {};

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];

    const rawDate = row[0]; // A
    const concept = row[1]; // B
    const value   = row[2]; // C
    const type    = row[3]; // D

    if (!concept || !(rawDate instanceof Date)) continue;

    // 🔹 Truncado de hora: YYYY-MM-DD 00:00
    const dateOnly = new Date(
      rawDate.getFullYear(),
      rawDate.getMonth(),
      rawDate.getDate()
    );

    if (!dict[concept]) {
      dict[concept] = {
        Type: type,
        Data: {
          Dates: [],
          Values: []
        }
      };
    }

    dict[concept].Data.Dates.push(dateOnly);
    dict[concept].Data.Values.push(value);
  }

  // --- 5) Log ---
  Logger.log("extractDatabaseDataAsDictAndLog_ (DATE-ONLY)");
  Logger.log(JSON.stringify(dict,null,2));
  return dict;
}

function calculateEMA_(conceptData, daysBack) {
  if (!conceptData || typeof conceptData !== "object" || !conceptData.Data) {
    throw new Error("Invalid conceptData object.");
  }

  const type = conceptData.Type;
  const Dates = conceptData.Data.Dates;
  const Values = conceptData.Data.Values;

  if (!Array.isArray(Dates) || !Array.isArray(Values)) {
    throw new Error("conceptData.Data.Dates and conceptData.Data.Values must be arrays.");
  }
  if (Dates.length !== Values.length) {
    throw new Error("Dates and Values length mismatch.");
  }
  if (Dates.length === 0) return null;

  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate()); // date-only
  const DAY_MS = 86400000;

  // --- Determine unique days present ---
  const uniqueDayKeys = new Set(
    Dates.map(d => d instanceof Date
      ? `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`
      : null
    ).filter(Boolean)
  );

  // 🔹 FALLBACK: not enough history → simple average
  if (uniqueDayKeys.size < 30) {
    return calculateAVG_(conceptData, null);
  }

  // --- EMA parameters (for num/date) ---
  const tau_default = (daysBack == null) ? Dates.length : daysBack;
  if (daysBack != null && (typeof daysBack !== "number" || daysBack <= 0)) {
    throw new Error("daysBack must be a positive number when provided.");
  }

  // --- BOOL PATH (modified: epoch + missing days treated as false) ---
  if (type === "bool") {
    // Window bounds:
    // - If daysBack is provided -> last N calendar days ending today (but not earlier than epoch)
    // - If daysBack is null/undefined -> from epoch to today
    let startDate;
    if (daysBack == null) {
      startDate = new Date(BOOL_EPOCH.getFullYear(), BOOL_EPOCH.getMonth(), BOOL_EPOCH.getDate());
    } else {
      startDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() - (daysBack - 1));
      if (startDate < BOOL_EPOCH) startDate = new Date(BOOL_EPOCH.getFullYear(), BOOL_EPOCH.getMonth(), BOOL_EPOCH.getDate());
    }

    if (startDate > today) return null;

    // Build daily ratios for days that have entries
    const dayMap = {}; // key -> {sum01, count}
    for (let i = 0; i < Values.length; i++) {
      const d = Dates[i];
      const v = Values[i];
      if (!(d instanceof Date)) continue;

      // Dates are already date-only from extractor, but keep safe:
      const d0 = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      if (d0 < startDate || d0 > today) continue;

      const key = `${d0.getFullYear()}-${d0.getMonth()}-${d0.getDate()}`;
      if (!dayMap[key]) dayMap[key] = { sum01: 0, count: 0 };

      const v01 = (v === true || v === 1) ? 1 : 0;
      dayMap[key].sum01 += v01;
      dayMap[key].count += 1;
    }

    // tau/alpha must be based on CALENDAR DAYS, not number of inputs
    const totalDays = Math.floor((today - startDate) / DAY_MS) + 1;
    const tau = (daysBack == null) ? totalDays : daysBack;
    const alpha = 1 - Math.exp(-1 / tau);

    // Iterate every day in the window, filling missing as 0 (false)
    let ema = null;
    for (let t = new Date(startDate); t <= today; t = new Date(t.getFullYear(), t.getMonth(), t.getDate() + 1)) {
      const key = `${t.getFullYear()}-${t.getMonth()}-${t.getDate()}`;
      const o = dayMap[key];
      const x = o ? (o.sum01 / o.count) : 0; // missing day => false

      if (ema === null) ema = x;
      else ema = alpha * x + (1 - alpha) * ema;
    }

    if (ema === null) return null;
    return Math.min(1, Math.max(0, ema));
  }

  // --- NUM / DATE (unchanged) ---
  const tau = tau_default;
  const alpha = 1 - Math.exp(-1 / tau);

  let ema = null;

  for (let i = 0; i < Values.length; i++) {
    const d = Dates[i];
    const v = Values[i];

    if (!(d instanceof Date)) continue;

    const ageDays = (now - d) / DAY_MS;
    if (daysBack != null && ageDays > daysBack) continue;

    if (typeof v !== "number" || !isFinite(v)) continue;

    ema = (ema === null) ? v : alpha * v + (1 - alpha) * ema;
  }

  if (ema === null) return null;

  if (type === "date") {
    let hours = Math.floor(ema);
    let minutes = Math.round((ema - hours) * 60);

    if (minutes >= 60) {
      hours += 1;
      minutes -= 60;
    }

    hours = Math.min(23, Math.max(0, hours));
    minutes = Math.min(59, Math.max(0, minutes));

    return new Date(
      now.getFullYear(),
      now.getMonth(),
      now.getDate(),
      hours,
      minutes,
      0,
      0
    );
  }

  return ema;
}



function calculateAVG_(conceptData, daysBack = 7) {
  if (!conceptData || !conceptData.Data) {
    throw new Error("Invalid conceptData object.");
  }

  const type = conceptData.Type;
  const Dates = conceptData.Data.Dates;
  const Values = conceptData.Data.Values;

  if (!Array.isArray(Dates) || !Array.isArray(Values)) {
    throw new Error("conceptData.Data.Dates and conceptData.Data.Values must be arrays.");
  }
  if (Dates.length !== Values.length) {
    throw new Error("Dates and Values length mismatch.");
  }
  if (Dates.length === 0) return null;

  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate()); // date-only
  const DAY_MS = 86400000;

  // Window bounds (inclusive) for non-bool and bool (bool has epoch rule below)
  let startDate = null;
  if (daysBack != null) {
    if (typeof daysBack !== "number" || daysBack <= 0) {
      throw new Error("daysBack must be a positive number, or null/undefined for all data.");
    }
    startDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() - (daysBack - 1));
  }

  // --- BOOL PATH (modified: epoch + missing days treated as false) ---
  if (type === "bool") {
    // Determine effective start:
    // - If daysBack is null => from epoch
    // - Else => last N days, but not earlier than epoch
    let effectiveStart;
    if (daysBack == null) {
      effectiveStart = new Date(BOOL_EPOCH.getFullYear(), BOOL_EPOCH.getMonth(), BOOL_EPOCH.getDate());
    } else {
      effectiveStart = startDate;
      if (effectiveStart < BOOL_EPOCH) effectiveStart = new Date(BOOL_EPOCH.getFullYear(), BOOL_EPOCH.getMonth(), BOOL_EPOCH.getDate());
    }

    if (effectiveStart > today) return null;

    // dayMap: daily true ratio where day has entries
    const dayMap = {}; // key -> {sum01, count}
    for (let i = 0; i < Values.length; i++) {
      const d = Dates[i];
      const v = Values[i];
      if (!(d instanceof Date)) continue;

      const d0 = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      if (d0 < effectiveStart || d0 > today) continue;

      const key = `${d0.getFullYear()}-${d0.getMonth()}-${d0.getDate()}`;
      if (!dayMap[key]) dayMap[key] = { sum01: 0, count: 0 };

      const v01 = (v === true || v === 1) ? 1 : 0;
      dayMap[key].sum01 += v01;
      dayMap[key].count += 1;
    }

    // Average across ALL days in the window (missing => 0)
    const totalDays = Math.floor((today - effectiveStart) / DAY_MS) + 1;

    let sumDaily = 0;
    for (let t = new Date(effectiveStart); t <= today; t = new Date(t.getFullYear(), t.getMonth(), t.getDate() + 1)) {
      const key = `${t.getFullYear()}-${t.getMonth()}-${t.getDate()}`;
      const o = dayMap[key];
      const x = o ? (o.sum01 / o.count) : 0; // missing day => false
      sumDaily += x;
    }

    const avgBool = sumDaily / totalDays;
    return Math.min(1, Math.max(0, avgBool));
  }

  // --- NON-BOOL PATH (unchanged) ---
  const dayMap = {}; // key -> {sum, count}

  for (let i = 0; i < Values.length; i++) {
    const d = Dates[i];
    const v = Values[i];

    if (!(d instanceof Date)) continue;

    if (startDate && d < startDate) continue;
    if (d > today) continue;

    const key = `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`;
    if (!dayMap[key]) dayMap[key] = { sum: 0, count: 0 };

    if (typeof v !== "number" || !isFinite(v)) continue;
    dayMap[key].sum += v;
    dayMap[key].count += 1;
  }

  const dayKeys = Object.keys(dayMap);
  if (dayKeys.length === 0) return null;

  let sumDailyMeans = 0;
  let daysWithData = 0;

  for (const k of dayKeys) {
    const o = dayMap[k];
    if (o.count <= 0) continue;
    const dailyMean = o.sum / o.count;
    sumDailyMeans += dailyMean;
    daysWithData += 1;
  }

  if (daysWithData === 0) return null;

  const avg = sumDailyMeans / daysWithData;

  if (type === "date") {
    let hours = Math.floor(avg);
    let minutes = Math.round((avg - hours) * 60);

    if (minutes >= 60) {
      hours += 1;
      minutes -= 60;
    }

    hours = Math.min(23, Math.max(0, hours));
    minutes = Math.min(59, Math.max(0, minutes));

    return new Date(today.getFullYear(), today.getMonth(), today.getDate(), hours, minutes, 0, 0);
  }

  return avg;
}


/************ CHART SETTINGS (baseline) ************/
/**
 * Configuración de render del gráfico (posición y escala) separada de los datos.
 * - position: celda ancla donde se inserta el gráfico (top-left)
 * - scale: factor de escala aplicado a BASE_WIDTH_PX y BASE_HEIGHT_PX
 * - baseSizePx: tamaño base antes de escalar
 */
const CHART_RENDER_CFG = {
  "Estado mental": {
        offset: 0,
        position: { row: 15, col: 6, offsetX: 0, offsetY: 0 }, // G4
        scale: 0.5,
        baseSizePx: { width: 801, height: 500 }
    },
};

function createDailyBarChartForConcept_(conceptName, conceptData, renderCfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shStats = ss.getSheetByName("Estadisticas");
  if (!shStats) throw new Error("No existe la hoja 'Estadisticas'.");

  if (!conceptName || typeof conceptName !== "string") {
    throw new Error("conceptName debe ser un string.");
  }
  if (!conceptData || typeof conceptData !== "object" || !conceptData.Data) {
    throw new Error("conceptData inválido (esperado inputDataDict[conceptName]).");
  }
  if (conceptData.Type !== "num") {
    throw new Error(`El concepto '${conceptName}' no es de tipo 'num' (es '${conceptData.Type}').`);
  }

  const Dates = conceptData.Data.Dates;
  const Values = conceptData.Data.Values;

  if (!Array.isArray(Dates) || !Array.isArray(Values) || Dates.length !== Values.length) {
    throw new Error("Datos inválidos: Dates/Values no son arrays o tienen longitudes distintas.");
  }

  // 1) Agregar por día: media diaria si hay múltiples entradas en el mismo día
  const dayMap = {}; // key -> {dateObj, sum, count}
  for (let i = 0; i < Dates.length; i++) {
    const d = Dates[i];
    const v = Values[i];

    if (!(d instanceof Date)) continue;
    if (typeof v !== "number" || !isFinite(v)) continue;

    const dateOnly = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    const key = `${dateOnly.getFullYear()}-${dateOnly.getMonth()}-${dateOnly.getDate()}`;

    if (!dayMap[key]) dayMap[key] = { dateObj: dateOnly, sum: 0, count: 0 };
    dayMap[key].sum += v;
    dayMap[key].count += 1;
  }

  const dailyRows = Object.values(dayMap)
    .sort((a, b) => a.dateObj - b.dateObj)
    .map(o => {
      const dd = String(o.dateObj.getDate()).padStart(2, "0");
      const mm = String(o.dateObj.getMonth() + 1).padStart(2, "0");
      return [`${dd}/${mm}`, o.sum / o.count];
    });

  if (dailyRows.length === 0) {
    throw new Error(`No hay datos numéricos válidos para '${conceptName}'.`);
  }

  // 2) Escribir tabla auxiliar en 'Estadisticas'
  //    - La tabla empieza en la fila 101
  //    - La columna depende de renderCfg.offset:
  //        offset=0 -> AA:AB
  //        offset=1 -> Y:Z
  //        offset=2 -> W:X
  //        ...
  //      (2 columnas a la izquierda por cada incremento)
  const TABLE_START_ROW = 101;

  const offsetN = (renderCfg && typeof renderCfg.offset === "number") ? Math.floor(renderCfg.offset) : 0;
  if (offsetN < 0) throw new Error("renderCfg.offset debe ser >= 0.");

  const BASE_PAIR_START_COL = 27; // AA = 27 (A=1)
  const pairStartCol = BASE_PAIR_START_COL - (offsetN * 2);
  if (pairStartCol < 1) {
    throw new Error(`renderCfg.offset demasiado grande (columna inicial < A). offset=${offsetN}`);
  }

  const startRow = TABLE_START_ROW;
  const startCol = pairStartCol;

  const table = [["Día", conceptName]].concat(dailyRows);

  shStats.getRange(startRow, startCol, table.length, 2).clearContent();
  shStats.getRange(startRow, startCol, table.length, 2).setValues(table);

  // 3) Eliminar gráficos previos con el mismo título para no acumular
  for (const ch of shStats.getCharts()) {
    const opts = ch.getOptions?.();
    const t = opts && opts.title ? String(opts.title) : "";
    if (t === conceptName) shStats.removeChart(ch);
  }

  // 4) Crear gráfico (ColumnChart) y colocarlo según renderCfg
  const dataRange = shStats.getRange(startRow, startCol, table.length, 2);

  const BASE_WIDTH_PX = (renderCfg && renderCfg.baseSizePx && renderCfg.baseSizePx.width) ? renderCfg.baseSizePx.width : 600;
  const BASE_HEIGHT_PX = (renderCfg && renderCfg.baseSizePx && renderCfg.baseSizePx.height) ? renderCfg.baseSizePx.height : 371;
  const SCALE = (renderCfg && typeof renderCfg.scale === "number") ? renderCfg.scale : 0.5;

  const chartWidth = Math.round(BASE_WIDTH_PX * SCALE);
  const chartHeight = Math.round(BASE_HEIGHT_PX * SCALE);

  const pos = (renderCfg && renderCfg.position) ? renderCfg.position : { row: 4, col: 7, offsetX: 0, offsetY: 0 };
  const posRow = (typeof pos.row === "number") ? pos.row : 4;
  const posCol = (typeof pos.col === "number") ? pos.col : 7;
  const offsetX = (typeof pos.offsetX === "number") ? pos.offsetX : 0;
  const offsetY = (typeof pos.offsetY === "number") ? pos.offsetY : 0;

  const chart = shStats.newChart()
    .asColumnChart()
    .addRange(dataRange)
    .setPosition(posRow, posCol, offsetX, offsetY)
    .setOption("titleTextStyle", {
      alignment: "center",
      fontSize: 20,
      fontName: "Times New Roman",
      color: "#000000"
    })
    .setOption("legend", { position: "none" })
    .setOption("width", chartWidth)
    .setOption("height", chartHeight)
    .setOption("annotations", {
      textStyle: {
        fontName: "Times New Roman",
        fontSize: 12
      }
    })
    .setOption("series", {
      0: { color: "#94849b" }
    })
    .build();

  shStats.insertChart(chart);

  Logger.log(
    `Chart creado para '${conceptName}' en 'Estadisticas'! ` +
    `(pos r${posRow} c${posCol}, size ${chartWidth}x${chartHeight}, ` +
    `data @ row ${startRow}, cols ${startCol}-${startCol + 1}, offset=${offsetN}).`
  );
}


function truncate2_(v) {
  if (typeof v !== "number" || !isFinite(v)) return 0;
  return Math.trunc(v * 100) / 100;
}





function clearAllChartsInEstadisticas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shStats = ss.getSheetByName("Estadisticas");
  if (!shStats) throw new Error("No existe la hoja 'Estadisticas'.");

  const charts = shStats.getCharts();
  Logger.log(`clearAllChartsInEstadisticas_: found ${charts.length} charts`);

  for (const ch of charts) {
    shStats.removeChart(ch);
  }

  Logger.log(`clearAllChartsInEstadisticas_: removed ${charts.length} charts`);
}


function fillEstadisticasFromInsightDict_(insightDict) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Estadisticas");
  if (!sh) throw new Error("No existe la hoja 'Estadisticas'.");

  if (!insightDict || typeof insightDict !== "object") {
    throw new Error("insightDict inválido.");
  }

  const START_ROW = 3;

  const maxRows = sh.getMaxRows() - START_ROW + 1;
  const colA = sh.getRange(START_ROW, 1, maxRows, 1).getValues(); // A

  const outB = [];
  const outC = [];
  const outD = [];

  const hourRowsRel = []; // filas (relativas) cuyo type es "hour"

  let nRowsUsed = 0;

  for (let i = 0; i < colA.length; i++) {
    const row = START_ROW + i;
    const raw = colA[i][0];

    // fin: primera celda vacía en A
    if (raw === "" || raw === null) break;

    nRowsUsed++;

    const concept = String(raw).trim();

    // Saltar secciones "< ... >" sin sobreescribir B/C/D
    if (concept.startsWith("<")) {
      outB.push([sh.getRange(row, 2).getValue()]);
      outC.push([sh.getRange(row, 3).getValue()]);
      outD.push([sh.getRange(row, 4).getValue()]);
      continue;
    }

    Logger.log(`Estadisticas row ${row}: concepto='${concept}'`);

    const entry = insightDict[concept];

    if (entry && typeof entry === "object") {
      const type = entry.type;

      let avgWeek  = (entry.avgWeek  == null) ? 0 : entry.avgWeek;
      let emaMonth = (entry.emaMonth == null) ? 0 : entry.emaMonth;
      let emaTotal = (entry.emaTotal == null) ? 0 : entry.emaTotal;

      if (type === "hour") {
        // Convert "hours as number" -> "fraction of day" for Sheets time formatting
        // Also round to nearest minute to avoid weird seconds
        const toDayFrac = (h) => {
          if (typeof h !== "number" || !isFinite(h)) return 0;
          const minutes = Math.round(h * 60);   // round to minute
          return (minutes / 60) / 24;           // minutes -> hours -> day fraction
        };

        outB.push([toDayFrac(avgWeek)]);
        outC.push([toDayFrac(emaMonth)]);
        outD.push([toDayFrac(emaTotal)]);

        hourRowsRel.push(nRowsUsed - 1);
    } else if (type === "bool") {
        // BOOL → percentage
        const toPct = (v) => truncate2_((typeof v === "number" && isFinite(v)) ? v * 100 : 0);

        outB.push([toPct(avgWeek)]);
        outC.push([toPct(emaMonth)]);
        outD.push([toPct(emaTotal)]);
    } else {
        // NUM
        outB.push([truncate2_(avgWeek)]);
        outC.push([truncate2_(emaMonth)]);
        outD.push([truncate2_(emaTotal)]);
        }
    } else {
      outB.push([0]);
      outC.push([0]);
      outD.push([0]);
    }
  }

  if (nRowsUsed === 0) {
    Logger.log("fillEstadisticasFromInsightDict_: no hay filas a procesar (A4 ya está vacío).");
    return;
  }

    // Escribir de golpe
    sh.getRange(START_ROW, 2, nRowsUsed, 1).setValues(outB); // B avgWeek
    sh.getRange(START_ROW, 3, nRowsUsed, 1).setValues(outC); // C emaMonth
    sh.getRange(START_ROW, 4, nRowsUsed, 1).setValues(outD); // D emaTotal

    // Formato porcentaje para filas bool
    for (let i = 0; i < nRowsUsed; i++) {
    const concept = String(colA[i][0]).trim();
    const entry = insightDict[concept];
    if (entry && entry.type === "bool") {
        sh.getRange(START_ROW + i, 2, 1, 3).setNumberFormat("0.##\"%\"");
    }
    }

  // Formato hora SOLO para filas type="hour" (B:C:D)
  for (const rel of hourRowsRel) {
    sh.getRange(START_ROW + rel, 2, 1, 3).setNumberFormat("hh:mm");
  }

  Logger.log(`fillEstadisticasFromInsightDict_: escrito B,C,D para ${nRowsUsed} filas (desde A4).`);
}





function roundToNearestHalf_(x) {
  const n = Number(x);
  if (!Number.isFinite(n)) return x;
  const eps = 1e-10;
  return Math.round((n + eps) * 2) / 2;
}

function parseTimeDisplay_(timeStr) {
  if (timeStr == null) return null;
  const s = String(timeStr).trim();
  if (s === "") return null;

  const m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return null;

  return {
    h: Number(m[1]),
    m: Number(m[2]),
    s: m[3] != null ? Number(m[3]) : 0
  };
}

function combineDateAndTime_(dateObj, timeStr) {
  if (!(dateObj instanceof Date)) return null;
  const t = parseTimeDisplay_(timeStr);
  if (!t) return null;

  const d = new Date(dateObj.getTime());
  d.setHours(t.h, t.m, t.s, 0);
  return d;
}

/**
 * MAIN
 */
function buildDatabaseDictFromWindow_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("database");
  if (!sh) throw new Error("No existe la hoja 'database'.");

  const startRow = Number(sh.getRange("B8").getValue()) + 1;
  const endRow   = Number(sh.getRange("D8").getValue()) - 1;

  if (!Number.isFinite(startRow) || !Number.isFinite(endRow)) {
    throw new Error("B8 y D8 deben contener números de fila válidos.");
  }
  if (endRow < startRow) return {};

  const numRows = endRow - startRow + 1;

  const values = sh.getRange(startRow, 1, numRows, 7).getValues();
  const timeDisplays = sh.getRange(startRow, 4, numRows, 2).getDisplayValues();

  const out = {};

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    const dateA   = row[0];                 // Col A
    const rawConcept = row[1];              // Col B
    const type    = row[2];                 // Col C
    const durRaw  = row[5];                 // Col F
    const done    = row[6];                 // Col G

    if (type == null || type === "" || rawConcept == null || rawConcept === "") continue;

    // 🔽 NORMALIZATION HERE
    const concept = String(rawConcept).toLowerCase();

    if (!out[type]) out[type] = {};

    if (!out[type][concept]) {
      out[type][concept] = {
        "start time": [],
        "end time": [],
        "duration": [],
        "wasDone": []
      };
    }

    const startDT = combineDateAndTime_(dateA, timeDisplays[i][0]);
    const endDT   = combineDateAndTime_(dateA, timeDisplays[i][1]);

    // Handle events crossing midnight
    if (startDT && endDT && endDT.getTime() < startDT.getTime()) {
      endDT.setDate(endDT.getDate() + 1);
    }

    out[type][concept]["start time"].push(startDT);
    out[type][concept]["end time"].push(endDT);
    out[type][concept]["duration"].push(roundToNearestHalf_(durRaw));
    out[type][concept]["wasDone"].push(done);

    // Alignment safety check
    const obj = out[type][concept];
    const L = obj["start time"].length;
    if (
      obj["end time"].length !== L ||
      obj["duration"].length !== L ||
      obj["wasDone"].length !== L
    ) {
      throw new Error(
        "Desalineación de arrays para type='" + type + "', concept='" + concept + "'."
      );
    }
  }
  Logger.log(JSON.stringify(out,null,2));
  return out;
}

function buildDoneHoursByDateAndType_(dbDict) {
  if (!dbDict || typeof dbDict !== "object") throw new Error("dbDict inválido.");

  const tz = Session.getScriptTimeZone();

  // ------------------------------------------------------------
  // CHANGE: date key is now day-month (dd-MM)
  // ------------------------------------------------------------
  function dateKey_(d) {
    return Utilities.formatDate(d, tz, "dd-MM");
  }

  function startOfDay_(d) {
    const x = new Date(d.getTime());
    x.setHours(0, 0, 0, 0);
    return x;
  }

  function addDays_(d, days) {
    const x = new Date(d.getTime());
    x.setDate(x.getDate() + days);
    return x;
  }

  function roundToNearestHalf_(x) {
    const n = Number(x);
    if (!Number.isFinite(n)) return x;
    const eps = 1e-10;
    return Math.round((n + eps) * 2) / 2;
  }

  // If an event ends exactly at midnight, treat it as not consuming the next day ([start, end))
  function adjustEndExclusive_(endDT) {
    if (!(endDT instanceof Date)) return endDT;
    if (
      endDT.getHours() === 0 &&
      endDT.getMinutes() === 0 &&
      endDT.getSeconds() === 0 &&
      endDT.getMilliseconds() === 0
    ) {
      return new Date(endDT.getTime() - 1);
    }
    return endDT;
  }

  const out = {};

  const typeKeys = Object.keys(dbDict);
  for (let t = 0; t < typeKeys.length; t++) {
    const type = typeKeys[t];
    if (type === "General") continue;

    const conceptsObj = dbDict[type];
    if (!conceptsObj || typeof conceptsObj !== "object") continue;

    const conceptKeys = Object.keys(conceptsObj);
    for (let c = 0; c < conceptKeys.length; c++) {
      const concept = conceptKeys[c];
      const node = conceptsObj[concept];
      if (!node) continue;

      const starts = node["start time"] || [];
      const ends   = node["end time"] || [];
      const dones  = node["wasDone"] || [];

      const n = Math.min(starts.length, ends.length, dones.length);
      for (let i = 0; i < n; i++) {
        if (dones[i] !== true) continue;

        const s0 = starts[i];
        const e0 = ends[i];
        if (!(s0 instanceof Date) || !(e0 instanceof Date)) continue;

        let s = new Date(s0.getTime());
        let e = new Date(e0.getTime());

        // Midnight crossing safety
        if (e.getTime() < s.getTime()) {
          e.setDate(e.getDate() + 1);
        }

        e = adjustEndExclusive_(e);

        let dayStart = startOfDay_(s);
        const lastDayStart = startOfDay_(e);

        while (dayStart.getTime() <= lastDayStart.getTime()) {
          const dayEnd = addDays_(dayStart, 1);

          const overlapMs = Math.max(
            0,
            Math.min(e.getTime(), dayEnd.getTime()) -
              Math.max(s.getTime(), dayStart.getTime())
          );

          if (overlapMs > 0) {
            const hours = overlapMs / (1000 * 60 * 60);
            const kDate = dateKey_(dayStart); // dd-MM

            if (!out[kDate]) out[kDate] = {};

            if (type === "Estudio") {
              if (!out[kDate]["Estudio"] || typeof out[kDate]["Estudio"] !== "object") {
                out[kDate]["Estudio"] = {};
              }
              if (!out[kDate]["Estudio"][concept]) out[kDate]["Estudio"][concept] = 0;
              out[kDate]["Estudio"][concept] += hours;
            } else {
              if (!out[kDate][type]) out[kDate][type] = 0;
              out[kDate][type] += hours;
            }
          }

          dayStart = dayEnd;
        }
      }
    }
  }

  // ------------------------------------------------------------
  // Round outputs to nearest 0.5
  // ------------------------------------------------------------
  const dates = Object.keys(out);
  for (let d = 0; d < dates.length; d++) {
    const dk = dates[d];
    const keys = Object.keys(out[dk]);

    for (let j = 0; j < keys.length; j++) {
      const k = keys[j];
      if (k === "Estudio" && out[dk][k] && typeof out[dk][k] === "object") {
        const subj = Object.keys(out[dk][k]);
        for (let s = 0; s < subj.length; s++) {
          const subk = subj[s];
          out[dk][k][subk] = roundToNearestHalf_(out[dk][k][subk]);
        }
      } else {
        out[dk][k] = roundToNearestHalf_(out[dk][k]);
      }
    }
  }

  Logger.log(JSON.stringify(out, null, 2));
  return out;
}




function renderDoneHoursStackedBarChart_(dailyDict, targetSheetName, anchorA1) {
  if (!dailyDict || typeof dailyDict !== "object") {
    throw new Error("dailyDict inválido.");
  }

  const ss = SpreadsheetApp.getActive();

  const targetName = targetSheetName || "MainDashboard";
  const anchor = anchorA1 || "B2";

  const targetSh = ss.getSheetByName(targetName);
  if (!targetSh) throw new Error("No existe la hoja destino: " + targetName);

  const dataSh = ss.getSheetByName("Estadisticas");
  if (!dataSh) throw new Error("No existe la hoja 'Estadisticas'.");

  const START_ROW = 101;
  const START_COL = 1; // column A

  // ------------------------------------------------------------------
  // Helpers: normalize label to dd-MM and provide sortable Date
  // ------------------------------------------------------------------
  function toDayMonthLabel_(x) {
    if (x instanceof Date) {
      const dd = String(x.getDate()).padStart(2, "0");
      const mm = String(x.getMonth() + 1).padStart(2, "0");
      return dd + "-" + mm;
    }

    const s = String(x).trim();

    // If already dd-MM
    let m = s.match(/^(\d{1,2})-(\d{1,2})$/);
    if (m) {
      const dd = String(m[1]).padStart(2, "0");
      const mm = String(m[2]).padStart(2, "0");
      return dd + "-" + mm;
    }

    // If yyyy-mm-dd or yyyy-m-d
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) {
      const mm = String(m[2]).padStart(2, "0");
      const dd = String(m[3]).padStart(2, "0");
      return dd + "-" + mm;
    }

    // Fallback: leave as-is (but you will see it)
    return s;
  }

  // Sort dd-MM correctly when crossing year boundary (e.g., 31-12 then 01-01),
  // using "today" as reference to infer the year.
  function sortKeyToDate_(keyDdMm) {
    const s = String(keyDdMm).trim();
    const m = s.match(/^(\d{2})-(\d{2})$/);
    if (!m) return new Date(0);

    const dd = Number(m[1]);
    const mm = Number(m[2]); // 1..12

    const now = new Date();
    const thisYear = now.getFullYear();
    const thisMonth = now.getMonth() + 1;

    // Heuristic:
    // if the month is "ahead" of current month, treat it as previous year (Dec when now is Jan, etc.)
    const year = (mm > thisMonth) ? (thisYear - 1) : thisYear;

    return new Date(year, mm - 1, dd, 0, 0, 0, 0);
  }

  // ------------------------------------------------------------------
  // 1) Collect all series names (normalizadas)
  // ------------------------------------------------------------------
  const rawDateKeys = Object.keys(dailyDict);

  // Normalize all date keys to dd-MM for sheet labels, but keep a mapping to fetch data
  const datePairs = rawDateKeys.map(k => {
    const label = toDayMonthLabel_(k);
    return { raw: k, label: label, sortDate: sortKeyToDate_(label) };
  });

  // Sort by inferred chronological date
  datePairs.sort((a, b) => a.sortDate - b.sortDate);

  const seriesSet = {};

  for (let i = 0; i < datePairs.length; i++) {
    const rawKey = datePairs[i].raw;
    const dayObj = dailyDict[rawKey] || {};

    Object.keys(dayObj).forEach(function (kRaw) {
      const k = String(kRaw).trim();
      if (!k) return;

      if (k === "Estudio" && typeof dayObj[kRaw] === "object") {
        const estudioObj = dayObj[kRaw];
        Object.keys(estudioObj).forEach(function (subkRaw) {
          const subk = String(subkRaw).trim();
          if (!subk) return;
          seriesSet["Estudio: " + subk] = true;
        });
      } else if (k !== "Estudio") {
        seriesSet[k] = true;
      }
    });
  }

  Logger.log("series set");
  Logger.log("series set");

  const seriesNames = Object.keys(seriesSet).sort(function (a, b) {
    const order = { "Proyectos": 1, "Tareas": 2 };
    const ao = order[a] || 100;
    const bo = order[b] || 100;
    if (ao !== bo) return ao - bo;
    return a.localeCompare(b);
  });

  // ------------------------------------------------------------------
  // 2) Build table [Date, series...]
  //    Writes dd-MM into column A (never yyyy-mm-dd)
  // ------------------------------------------------------------------
  const header = ["Date"].concat(seriesNames);
  const table = [header];

  for (let i = 0; i < datePairs.length; i++) {
    const rawKey = datePairs[i].raw;
    const dLabel = datePairs[i].label;

    const dayObj = dailyDict[rawKey] || {};
    const row = [dLabel];

    for (let s = 0; s < seriesNames.length; s++) {
      const series = seriesNames[s];
      let v = 0;

      if (series === "Proyectos") {
        v = Number(dayObj["Proyectos"] || 0);
      } else if (series === "Tareas") {
        v = Number(dayObj["Tareas"] || 0);
      } else if (series.startsWith("Estudio: ")) {
        const subk = series.slice(9);
        v = (dayObj.Estudio && dayObj.Estudio[subk]) || 0;
      } else {
        v = Number(dayObj[series] || 0);
      }

      row.push(Number.isFinite(v) ? v : 0);
    }

    table.push(row);
  }

  // ------------------------------------------------------------------
  // 2.5) Remove series columns that are entirely zero
  // ------------------------------------------------------------------
  const keepColIdx = [0]; // siempre mantener Date

  for (let c = 1; c < table[0].length; c++) {
    const headerName = String(table[0][c]).trim();
    if (!headerName) continue;

    let hasNonZero = false;
    for (let r = 1; r < table.length; r++) {
      if (Number(table[r][c] || 0) !== 0) {
        hasNonZero = true;
        break;
      }
    }

    if (hasNonZero) keepColIdx.push(c);
  }

  const filteredTable = [];
  for (let r = 0; r < table.length; r++) {
    const newRow = [];
    for (let i = 0; i < keepColIdx.length; i++) {
      newRow.push(table[r][keepColIdx[i]]);
    }
    filteredTable.push(newRow);
  }

  // sustituimos table por la versión filtrada
  table.length = 0;
  Array.prototype.push.apply(table, filteredTable);

  // ------------------------------------------------------------------
  // 3) Write table into Estadisticas!A101
  // ------------------------------------------------------------------
  const numRows = table.length;
  const numCols = table[0].length;

  dataSh
    .getRange(START_ROW, START_COL, numRows, numCols)
    .clearContent()
    .setValues(table);

  // ------------------------------------------------------------------
  // 4) Remove existing charts on target sheet
  // ------------------------------------------------------------------
  const existingCharts = targetSh.getCharts();
  for (let i = 0; i < existingCharts.length; i++) {
    targetSh.removeChart(existingCharts[i]);
  }

  // ------------------------------------------------------------------
  // 4.5) Build per-series color options
  //    - Estudio:*  -> blue
  //    - Proyecto   -> yellow
  //    - Tareas     -> green
  // ------------------------------------------------------------------
  const BLUE = "#4285F4";
  const YELLOW = "#FBBC05";
  const GREEN = "#34A853";

  const seriesOptions = {};
  for (let c = 1; c < table[0].length; c++) {
    const name = String(table[0][c]).trim();
    if (!name) continue;

    const seriesIdx = c - 1; // chart series are 0-based, excluding the domain column

    if (name === "Proyectos") {
      seriesOptions[seriesIdx] = { color: YELLOW };
    } else if (name === "Tareas") {
      seriesOptions[seriesIdx] = { color: GREEN };
    } else if (name.startsWith("Estudio: ")) {
      seriesOptions[seriesIdx] = { color: BLUE };
    }
  }

  // ------------------------------------------------------------------
  // 5) Build stacked column chart
  // ------------------------------------------------------------------
  Logger.log("START ROW ASDADS");
  Logger.log(START_ROW);

  const dataRange = dataSh.getRange(
    START_ROW,
    START_COL,
    numRows,
    numCols
  );

  const CHART_RENDER_CFG_2 = {
    position: { row: 3, col: 6, offsetX: 0, offsetY: 0 }, // G4
    scale: 0.5,
    baseSizePx: { width: 801, height: 500 }
  };

  const cfg = CHART_RENDER_CFG_2;
  const posRow  = cfg.position.row;
  const posCol  = cfg.position.col;
  const offsetX = cfg.position.offsetX || 0;
  const offsetY = cfg.position.offsetY || 0;

  const chartWidth  = Math.round(cfg.baseSizePx.width  * cfg.scale);
  const chartHeight = Math.round(cfg.baseSizePx.height * cfg.scale);

  const chart = targetSh.newChart()
    .asColumnChart()
    .addRange(dataRange)
    .setNumHeaders(1)
    .setOption("isStacked", true)
    .setOption("legend", {
      position: "bottom",
      textStyle: {
        fontSize: 10   // default is ~12
      }
    })
    .setOption("series", seriesOptions)
    .setOption("width", chartWidth)
    .setOption("height", chartHeight)
    .setPosition(posRow, posCol, offsetX, offsetY)
    .build();

  targetSh.insertChart(chart);
}





function mainTest(){

    clearAllChartsInEstadisticas_();
    const calendarCompletionDict = buildDatabaseDictFromWindow_();
    const byDay = buildDoneHoursByDateAndType_(calendarCompletionDict);
    renderDoneHoursStackedBarChart_(byDay, "Estadisticas", "I8", "_DoneHoursChartData");


    // Sacamos los datos de la sección de inputs
    const inputDataDict = extractDatabaseDataAsDictAndLog_()
    let listOfConcepts = Object.keys(inputDataDict);
    let insightDict = {};
    for (const concept of listOfConcepts){
        insightDict[concept] = {
            "type":inputDataDict[concept]["Type"],
            "emaTotal":calculateEMA_(inputDataDict[concept]),
            "emaMonth":calculateEMA_(inputDataDict[concept],30),
            "avgWeek":calculateAVG_(inputDataDict[concept])
        };
    }
    Logger.log(JSON.stringify(insightDict,null,2));
    createDailyBarChartForConcept_("Estado mental",inputDataDict["Estado mental"],CHART_RENDER_CFG["Estado mental"]);
    fillEstadisticasFromInsightDict_(insightDict);
}
