/************ CONFIG ************/
const DATABASE_SHEET_NAME  = "database";      // storage sheet (same file)

// Each dashboard sheet config lives here (same code handles both)
const DASHBOARDS = [
  {
    sheetName: "Proyectos",
    dropdownCellA1: "A1",
    startRow: 6,
    endRow: 44,
    startCol: 3,                 // C
    numberOfColumnsToBackUp: 13, // C..O
    // Mapping table for "Configuración proyectos" in database: D:E:F from row 5 down
    tableFirstDataRow: 5,
    mapColConcept: 5,  // D
    mapColStart: 6,    // E
    mapColEnd: 7       // F
  },
  {
    sheetName: "EstudioBien",
    dropdownCellA1: "A1",
    startRow: 3,
    endRow: 44,
    startCol: 3,                 // C
    numberOfColumnsToBackUp: 13, // C..O (keep same for now)
    // Mapping table for "Configuración estudios" in database: I:J:K from row 5 down
    tableFirstDataRow: 5,
    mapColConcept: 10,  // I
    mapColStart: 11,   // J
    mapColEnd: 12      // K
  }
];

// DB blocks always start at column C (where you store the dashboard blocks)
const DB_BLOCK_START_COL  = 3;  // C

/************ TRIGGER ************/

/************ MAIN CHUNK ************/
/**
 * Swaps the dashboard block to/from the database, keyed by the dropdown value.
 * Works for multiple dashboard sheets with different configs.
 *
 * @param {Event} e
 * @param {Object} cfg
 */
function swapDashboardFromDatabase_(e, cfg) {
  const ss = e.source;

  const newKey = (e.value || "").trim();
  const oldKey = (e.oldValue || "").trim();
  if (!newKey) return;

  const dashSheet = ss.getSheetByName(cfg.sheetName);
  const dbSheet   = ss.getSheetByName(DATABASE_SHEET_NAME);

  if (!dashSheet || !dbSheet) {
    throw new Error("Missing required sheet(s). Check sheet names in CONFIG.");
  }

  const maxRows = cfg.endRow - cfg.startRow + 1;

  // 1) Save current dashboard block to DB at oldKey location
  if (oldKey) {
    const oldLoc = getLocationFromDb_(dbSheet, oldKey, cfg);
    if (!oldLoc) throw new Error("No location found for old key: " + oldKey + " (sheet " + cfg.sheetName + ")");

    const dashBlock = readDashboardBlock_(dashSheet, cfg);
    writeBlockToDb_(dbSheet, oldLoc.startRow, oldLoc.endRow, dashBlock.values, dashBlock.formulas, cfg.numberOfColumnsToBackUp);
  }

  // 2) Clear dashboard block
  clearDashboardBlock_(dashSheet, cfg);

  // 3) Load newKey block from DB into dashboard
  const newLoc = getLocationFromDb_(dbSheet, newKey, cfg);
  if (!newLoc) throw new Error("No location found for new key: " + newKey + " (sheet " + cfg.sheetName + ")");

  const newBlock = readBlockFromDb_(dbSheet, newLoc.startRow, newLoc.endRow, cfg.numberOfColumnsToBackUp);

  if (newBlock.values.length > 0) {
    const numRows = Math.min(newBlock.values.length, maxRows);

    // Write all values (this restores normal cells correctly)
    dashSheet
      .getRange(cfg.startRow, cfg.startCol, numRows, cfg.numberOfColumnsToBackUp)
      .setValues(newBlock.values.slice(0, numRows));

    // IMPORTANT FIX:
    // Do NOT call setFormulas() on the whole range with blanks, because "" formulas clear values.
    // Instead, apply formulas only to the cells that actually have a formula.
    setFormulasSparse_(dashSheet, cfg.startRow, cfg.startCol, newBlock.formulas.slice(0, numRows));
  }
}

/************ HELPERS ************/

// Finds key row in DATABASE sheet (per-dashboard mapping table) and returns {startRow, endRow}
function getLocationFromDb_(dbSheet, keyName, cfg) {
  const lastRow = dbSheet.getLastRow();
  if (lastRow < cfg.tableFirstDataRow) return null;

  const numRows = lastRow - cfg.tableFirstDataRow + 1;

  // Read Concepto/Inicio/Final from the configured columns
  const values = dbSheet
    .getRange(cfg.tableFirstDataRow, cfg.mapColConcept, numRows, 3)
    .getValues();

  for (let i = 0; i < values.length; i++) {
    const concepto = String(values[i][0] || "").trim();

    // stop at first blank concepto (end of table)
    if (!concepto) break;

    if (concepto === keyName) {
      const startRow = Number(values[i][1]);
      const endRow   = Number(values[i][2]);
      if (!startRow || !endRow || endRow < startRow) return null;
      return { startRow, endRow };
    }
  }
  return null;
}

// Reads dashboard block strictly from cfg.startCol/cfg.startRow for cfg.numberOfColumnsToBackUp width
// Captures both values and formulas
function readDashboardBlock_(dashSheet, cfg) {
  const numRows = cfg.endRow - cfg.startRow + 1;

  const range = dashSheet.getRange(cfg.startRow, cfg.startCol, numRows, cfg.numberOfColumnsToBackUp);

  const values = range.getValues();
  const formulas = range.getFormulas(); // "" where there is no formula

  // For formula cells, blank the value so DB doesn't store calculated results as static values
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (formulas[r][c]) values[r][c] = "";
    }
  }

  return { values, formulas };
}

// Clears dashboard block strictly from cfg.startCol/cfg.startRow for cfg.numberOfColumnsToBackUp width
function clearDashboardBlock_(dashSheet, cfg) {
  const numRows = cfg.endRow - cfg.startRow + 1;

  dashSheet
    .getRange(cfg.startRow, cfg.startCol, numRows, cfg.numberOfColumnsToBackUp)
    .clearContent();
}

// Writes the dashboard block into DB between [startRow..endRow] but ONLY starting at column C for "cols" width.
// Writes values normally, and applies formulas ONLY where there are formulas (sparse), so non-formula values are not destroyed.
function writeBlockToDb_(dbSheet, startRow, endRow, values, formulas, cols) {
  const targetRows = endRow - startRow + 1;
  if (targetRows <= 0) return;

  const normalizedValues = [];
  const normalizedFormulas = [];

  for (let r = 0; r < targetRows; r++) {
    const outV = new Array(cols).fill("");
    const outF = new Array(cols).fill("");

    if (values[r]) {
      for (let c = 0; c < cols; c++) outV[c] = values[r][c] ?? "";
    }
    if (formulas[r]) {
      for (let c = 0; c < cols; c++) outF[c] = formulas[r][c] ?? "";
    }

    normalizedValues.push(outV);
    normalizedFormulas.push(outF);
  }

  // 1) write values
  dbSheet
    .getRange(startRow, DB_BLOCK_START_COL, targetRows, cols)
    .setValues(normalizedValues);

  // 2) apply formulas sparsely (CRITICAL FIX)
  setFormulasSparse_(dbSheet, startRow, DB_BLOCK_START_COL, normalizedFormulas);
}

// Reads block from DB between [startRow..endRow] starting at column C for "cols" width, trims trailing empty rows.
// Returns both values and formulas; trailing-trim considers both.
function readBlockFromDb_(dbSheet, startRow, endRow, cols) {
  const numRows = endRow - startRow + 1;
  if (numRows <= 0) return { values: [], formulas: [] };

  const range = dbSheet.getRange(startRow, DB_BLOCK_START_COL, numRows, cols);

  const values = range.getValues();
  const formulas = range.getFormulas(); // "" where not a formula

  let last = -1;
  for (let i = 0; i < numRows; i++) {
    const rowHasValue = values[i].some(v => v !== "" && v !== null);
    const rowHasFormula = formulas[i].some(f => f !== "" && f !== null);
    if (rowHasValue || rowHasFormula) last = i;
  }

  if (last === -1) return { values: [], formulas: [] };

  // For formula cells, blank the value so restore doesn't paste calculated results first
  const outValues = values.slice(0, last + 1).map((row, rIdx) => {
    const rr = row.slice();
    for (let c = 0; c < cols; c++) {
      if (formulas[rIdx][c]) rr[c] = "";
    }
    return rr;
  });

  const outFormulas = formulas.slice(0, last + 1);

  return { values: outValues, formulas: outFormulas };
}

/**
 * Applies formulas to a sheet ONLY where formulas[r][c] is non-empty.
 * This avoids the Google Sheets behavior where setFormulas("") clears values.
 *
 * @param {Sheet} sheet
 * @param {number} startRow
 * @param {number} startCol
 * @param {string[][]} formulas
 */
function setFormulasSparse_(sheet, startRow, startCol, formulas) {
  for (let r = 0; r < formulas.length; r++) {
    const row = formulas[r];
    if (!row) continue;

    for (let c = 0; c < row.length; c++) {
      const f = row[c];
      if (f !== "" && f !== null) {
        sheet.getRange(startRow + r, startCol + c).setFormula(f);
      }
    }
  }
}
