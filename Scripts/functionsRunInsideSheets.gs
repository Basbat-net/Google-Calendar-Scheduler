/******************** CONFIG / TESTING ********************/

// Hour offset applied ONLY in "from now" mode.
// Use 0 for real behaviour.
// Examples: +2, -1, -3.5
const TEST_HOUR_OFFSET = 0;

// Calendar sheet and grid definition
const CAL_SHEET_NAME = "Calendario";

// Your grid (as you described):
// - Row 1: headers (Monday..Sunday)
// - Row 2: 09:30
// - Row 49: 09:00
const GRID_HEADER_ROW = 1;
const GRID_FIRST_TIME_ROW = 2;
const GRID_LAST_TIME_ROW = 49;

// Columns:
// A = time text
// B..O = week grid in pairs
const TIME_COL = 1;     // A
const FIRST_WEEK_COL = 2;  // B
const LAST_WEEK_COL = 15;  // O


/******************** PUBLIC SPREADSHEET FUNCTIONS ********************/

/**
 * Counts tagged events across the entire week (B:O in pairs).
 * Usage:
 *   =COUNT_EVENTS_WEEK("Tareas")
 *   =COUNT_EVENTS_WEEK(A1)
 */
function COUNT_EVENTS_WEEK(tag) {
  return countTaggedSlots_WholeWeek_BtoO_(tag);
}

/**
 * Counts tagged events from "now" (rounded down to :00 / :30)
 * until end of the week.
 * Usage:
 *   =COUNT_EVENTS_FROM_NOW("Tareas")
 *   =COUNT_EVENTS_FROM_NOW(A1)
 */
function COUNT_EVENTS_FROM_NOW(tag) {
  return countTaggedSlots_WholeWeek_BtoO_(tag, null);
}


/******************** CORE LOGIC ********************/

/**
 * Main entry point.
 * - If mode is undefined → count whole week
 * - If mode === null     → count from now → end of week
 *
 * Optimized: one bulk read for values + one bulk read for merges.
 */
function countTaggedSlots_WholeWeek_BtoO_(tag, mode) {
  const tagText = normalizeTag_(tag);
  if (!tagText) return 0;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CAL_SHEET_NAME);
  if (!sh) throw new Error(`Sheet not found: ${CAL_SHEET_NAME}`);

  // We only ever scan the known grid rows/cols (fast, deterministic)
  const numRows = GRID_LAST_TIME_ROW;   // 49
  const numCols = LAST_WEEK_COL;        // 15 (A..O)
  const gridRange = sh.getRange(1, 1, numRows, numCols);

  // Bulk values (includes header row and time col)
  const values = gridRange.getDisplayValues(); // [rowIndex0][colIndex0]

  // Bulk merged ranges only within the week grid area (B..O, rows 2..49)
  const weekGridRange = sh.getRange(GRID_FIRST_TIME_ROW, FIRST_WEEK_COL,
                                   GRID_LAST_TIME_ROW - GRID_FIRST_TIME_ROW + 1,
                                   LAST_WEEK_COL - FIRST_WEEK_COL + 1);
  const mergedRanges = weekGridRange.getMergedRanges();

  // Build fast merge maps for O(1) lookup per cell.
  const mergeMaps = buildMergeMaps_(mergedRanges, values);

  // Determine start point
  let startCol = FIRST_WEEK_COL;       // B
  let startRow = GRID_FIRST_TIME_ROW;  // 2

  if (mode === null) {
    const start = getStartCellFromNow_();
    startCol = start.startCol;
    startRow = start.startRow;
  }

  let total = 0;

  // Iterate pairs from startCol to end (B..O in steps of 2)
  for (let col = startCol; col <= LAST_WEEK_COL; col += 2) {
    const pairStartRow = (col === startCol) ? startRow : GRID_FIRST_TIME_ROW;

    total += countTaggedSlots_ColumnPair_Fast_(values, mergeMaps, {
      col1: col,
      col2: col + 1,
      startRow: pairStartRow,
      tagText: tagText
    });
  }

  return total;
}


/******************** FAST COUNTING (IN-MEMORY) ********************/

/**
 * Counts a single pair using:
 * - values[][] already loaded
 * - merge maps already built
 *
 * Preserves your original semantics:
 * - Each column is counted independently with its own cursor
 * - Merged cells count as their merged height
 * - If starting inside a merged block, counts remaining part only
 */
function countTaggedSlots_ColumnPair_Fast_(values, mergeMaps, cfg) {
  const col1 = cfg.col1;
  const col2 = cfg.col2;
  const startRow = cfg.startRow;
  const tagText = cfg.tagText;

  let count = 0;

  // Normalize each cursor if starting inside a merge
  const n1 = normalizeStartCursorFast_(values, mergeMaps, startRow, col1, tagText);
  const n2 = normalizeStartCursorFast_(values, mergeMaps, startRow, col2, tagText);

  let r1 = n1.row;
  let r2 = n2.row;
  count += n1.add + n2.add;

  // Time awareness (kept; you said it will be useful later)
  // values[r-1][0] is the time text in column A.

  // Main scan runs only over grid rows (2..49)
  for (let rTime = startRow; rTime <= GRID_LAST_TIME_ROW; rTime++) {
    if (r1 > GRID_LAST_TIME_ROW && r2 > GRID_LAST_TIME_ROW) break;

    // Current time text available if needed later:
    const timeStr = String(values[rTime - 1][TIME_COL - 1] || "").trim();
    void timeStr;

    if (r1 === rTime && r1 <= GRID_LAST_TIME_ROW) {
      const adv1 = advanceAtCellFast_(values, mergeMaps, r1, col1, tagText);
      count += adv1.add;
      r1 += adv1.advance;
    }

    if (r2 === rTime && r2 <= GRID_LAST_TIME_ROW) {
      const adv2 = advanceAtCellFast_(values, mergeMaps, r2, col2, tagText);
      count += adv2.add;
      r2 += adv2.advance;
    }
  }

  return count;
}


/******************** MERGE MAPS ********************/

/**
 * Builds:
 * - anchorHeight[key] = merged height, only for "anchor" cells (top row per column in that merge)
 * - insideBottom[key] = bottom row of the merge for any cell inside a merge (including anchor)
 * - insideAnchorRow[key] = anchor row (top row) for that cell (per column)
 *
 * Key format: "r,c" where r and c are 1-based sheet coordinates.
 *
 * Important:
 * - We treat merges as vertical occupancy per column.
 * - If a merged range spans multiple columns, we create a per-column anchor at (topRow, col)
 *   so each column’s cursor logic remains consistent and does not “miss” the merge.
 */
function buildMergeMaps_(mergedRanges, values) {
  const anchorHeight = Object.create(null);
  const insideBottom = Object.create(null);
  const insideAnchorRow = Object.create(null);

  for (let i = 0; i < mergedRanges.length; i++) {
    const mr = mergedRanges[i];

    // mr is within weekGridRange; convert to absolute sheet coords
    const topRow = mr.getRow();
    const leftCol = mr.getColumn();
    const numRows = mr.getNumRows();
    const numCols = mr.getNumColumns();

    const bottomRow = topRow + numRows - 1;

    // For each column spanned by this merge, create a per-column anchor at (topRow, col)
    for (let dc = 0; dc < numCols; dc++) {
      const col = leftCol + dc;

      // Anchor
      anchorHeight[key_(topRow, col)] = numRows;

      // Mark all cells in that column as inside this merge (for mid-merge start detection)
      for (let dr = 0; dr < numRows; dr++) {
        const row = topRow + dr;
        insideBottom[key_(row, col)] = bottomRow;
        insideAnchorRow[key_(row, col)] = topRow;
      }
    }
  }

  return { anchorHeight, insideBottom, insideAnchorRow };
}

function key_(r, c) {
  return r + "," + c;
}


/******************** FAST CELL ADVANCE ********************/

/**
 * Handles normal step at (row,col):
 * - If (row,col) is an anchor of a merged block: add mergedHeight if tag matches, advance mergedHeight
 * - Else if inside merge but not anchor: add 0, advance 1
 * - Else (not merged): add 1 if tag matches, advance 1
 */
function advanceAtCellFast_(values, mergeMaps, row, col, tagText) {
  const k = key_(row, col);

  const h = mergeMaps.anchorHeight[k];
  if (h) {
    const v = String(values[row - 1][col - 1] || "");
    const containsTag = v.indexOf(tagText) !== -1;
    return { add: containsTag ? h : 0, advance: h };
  }

  // If inside a merge but not anchor, do not double count
  if (mergeMaps.insideBottom[k]) {
    return { add: 0, advance: 1 };
  }

  // Normal cell
  const v = String(values[row - 1][col - 1] || "");
  return { add: (v.indexOf(tagText) !== -1) ? 1 : 0, advance: 1 };
}

/**
 * Fix for starting inside a merged event:
 * If startRow is inside a merge but not on its anchor row,
 * count only the remaining part and jump cursor to bottom+1.
 */
function normalizeStartCursorFast_(values, mergeMaps, startRow, col, tagText) {
  const k = key_(startRow, col);
  const bottom = mergeMaps.insideBottom[k];
  const anchorRow = mergeMaps.insideAnchorRow[k];

  if (!bottom || !anchorRow) {
    return { row: startRow, add: 0 };
  }

  // If we are on anchor row, let normal logic handle it
  if (startRow === anchorRow) {
    return { row: startRow, add: 0 };
  }

  // We are inside merge; check tag on the anchor cell for this column
  const v = String(values[anchorRow - 1][col - 1] || "");
  const containsTag = v.indexOf(tagText) !== -1;

  const remaining = bottom - startRow + 1;

  return {
    row: bottom + 1,
    add: containsTag ? remaining : 0
  };
}


/******************** START CELL FROM NOW ********************/

/**
 * Computes start column + row from current time for "from now" mode.
 * - Applies TEST_HOUR_OFFSET
 * - Rounds DOWN to :00 or :30
 * - Maps day-of-week to B/D/F/.../N
 * - Maps time to row within 09:30..09:00 grid (rows 2..49)
 */
function getStartCellFromNow_() {
  const now = new Date();
  const shifted = new Date(now.getTime() + TEST_HOUR_OFFSET * 3600000);

  // Round DOWN to :00 or :30
  const rounded = new Date(shifted.getTime());
  rounded.setMinutes(rounded.getMinutes() - (rounded.getMinutes() % 30), 0, 0);

  // Day mapping: JS 0=Sun..6=Sat -> weekIndex Mon=0..Sun=6
  const jsDay = rounded.getDay();
  const weekIndex = (jsDay === 0) ? 6 : (jsDay - 1);

  const startCol = FIRST_WEEK_COL + weekIndex * 2;

  // Time mapping
  const minutesNow = rounded.getHours() * 60 + rounded.getMinutes();
  const gridStart = 9 * 60 + 30; // 09:30

  // grid spans 47 steps (row 2 -> 49) => 47*30 = 1410 minutes = 23h30
  const maxSteps = GRID_LAST_TIME_ROW - GRID_FIRST_TIME_ROW; // 47
  const maxRange = maxSteps * 30; // 1410

  // If before 09:30, clamp to start
  if (minutesNow <= gridStart) {
    return { startCol, startRow: GRID_FIRST_TIME_ROW };
  }

  // delta from 09:30, if after midnight then wrap +1440
  let delta = minutesNow - gridStart;
  if (delta < 0) delta += 1440;

  // clamp to grid range
  if (delta > maxRange) delta = maxRange;

  const steps = Math.floor(delta / 30);
  const startRow = GRID_FIRST_TIME_ROW + steps;

  return { startCol, startRow };
}


/******************** UTILITIES ********************/

/**
 * Normalizes "Tareas" → "[Tareas]"
 */
function normalizeTag_(tag) {
  if (!tag) return "";
  const t = String(tag).trim();
  if (!t) return "";
  return (t.startsWith("[") && t.endsWith("]")) ? t : `[${t}]`;
}


/**
 * Returns the last row of the merged region that contains (startRow, colLetter),
 * or startRow if the cell is not merged.
 *
 * @param {string} sheetName Name of the sheet
 * @param {string} colLetter Column letter(s), e.g. "A", "AA"
 * @param {number} startRow  1-based row index
 * @return {number}
 */
function getMergeEnd(sheetName, colLetter, startRow) {
  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(sheetName);

  if (!sheet) {
    throw new Error("Sheet not found");
  }

  const col = columnLetterToNumber(colLetter);
  const cell = sheet.getRange(startRow, col);

  if (!cell.isPartOfMerge()) return startRow;

  const merged = cell.getMergedRanges();
  if (!merged.length) return startRow;

  // ✅ FIX: escoger el merge que CONTIENE exactamente esta celda,
  // no asumir que merged[0] es el correcto.
  let r = null;
  for (let i = 0; i < merged.length; i++) {
    const m = merged[i];
    const r0 = m.getRow();
    const c0 = m.getColumn();
    const r1 = r0 + m.getNumRows() - 1;
    const c1 = c0 + m.getNumColumns() - 1;

    if (startRow >= r0 && startRow <= r1 && col >= c0 && col <= c1) {
      r = m;
      break;
    }
  }

  // Si por lo que sea no encontramos uno que contenga la celda, fallback seguro
  if (!r) return startRow;

  return r.getRow() + r.getNumRows() - 1;
}

/**
 * Converts column letters to 1-based index.
 */
function columnLetterToNumber(letter) {
  let col = 0;
  for (const c of letter.toUpperCase()) {
    col = col * 26 + (c.charCodeAt(0) - 64);
  }
  return col;
}
