function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️Manual Functions🛠️')
    .addItem('Submit All', 'submitAll')
    .addSeparator()
    .addItem('Submit Weight', 'submitWeight')
    .addItem('Submit Body Fat', 'submitBodyFat')
    .addItem('Submit Nutrition', 'submitNutrition')
    .addItem('Submit Water', 'submitWater')
    .addItem('Submit Supplements', 'submitSupplements')
    .addItem('Submit Skincare', 'submitSkincare')
    .addItem('Submit Vyvanse', 'submitVyvanse')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;

  const sh        = e.range.getSheet();
  const sheetName = sh.getName();
  const r         = e.range.getRow();
  const c1        = e.range.getColumn();
  const c2        = c1 + e.range.getNumColumns() - 1;

  // A) Food Nutrition Reference: sort when edit touches col B or V (row >= 3)
  if (sheetName === 'Food Nutrition Reference') {
    const rows   = e.range.getNumRows();
    const cols   = e.range.getNumColumns();
    const bottom = r + rows - 1;
    const right  = c1 + cols - 1;

    const touchesB = bottom >= 3 && c1 <= 2  && right >= 2;   // intersects column B
    const touchesV = bottom >= 3 && c1 <= 22 && right >= 22;  // intersects column V

    if (touchesB || touchesV) {
      try { e.source.toast(`Ref edit @ r${r} c${c1}-${right} | B=${touchesB} V=${touchesV}`, 'onEdit', 3); } catch (_) {}

      try {
        sortReferenceTables_(); // sorts B:T and V:AN
        const helper = e.source.getSheetByName('_AutocompleteHelper');
        if (helper) writeHelperBlocks_(helper, getLiveRefNames_());
        SpreadsheetApp.flush();
        try { e.source.toast('Sorted & helper refreshed', 'onEdit', 2); } catch (_){}
      } catch (err) {
        try { e.source.toast('Sort failed: ' + err, 'onEdit ERR', 6); } catch (_){}
        throw err;
      }
    }
    return;
  }

  // B) Data Tables: if L4:AC is edited, mirror that row to Inputs!I17 + N17:AD17
  if (sheetName === 'Data Tables') {
    const TABLE_TOP = 4, LEFT = 12, RIGHT = 29; // L..AC
    const bottom = r + e.range.getNumRows() - 1;

    const intersectsNutrition = bottom >= TABLE_TOP && c2 >= LEFT && c1 <= RIGHT;
    if (intersectsNutrition) {
      const targetRow = Math.max(TABLE_TOP, r); // mirror the top row of the edited block
      mirrorNutritionPreviewFromRow_(targetRow);
    }
    return;
  }

  // C) Inputs sheet behavior
  if (sheetName !== 'Inputs') return;

  // 1) Auto-fill today’s date (no time) in B4, E4, I4:I8, AF4, AI4, AS4, AY4 if blank
  fillTodayDateIfBlank_(sh);

  // 2) If typing in J4:J13, update that row's helper block with fuzzy matches
  if (
    r >= 4 && r <= 13 &&
    c1 === 10 &&
    e.range.getNumRows() === 1 &&
    e.range.getNumColumns() === 1
  ) {
    const typed = (e.value || '').toString();
    updateNutritionSuggestionsBlock_(r, typed);
  }
}

 // Submit All function for all Inputs
function submitAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();
  const results = [];

  try { lock.waitLock(30000); } catch (_) {}

  const tasks = [
    ['Weight',        submitWeight],
    ['Body Fat',      submitBodyFat],
    ['Nutrition',     submitNutrition],
    ['Water',         submitWater],
    ['Supplements',   submitSupplements],
    ['Skincare',      submitSkincare],
    ['Vyvanse',       submitVyvanse],
  ];

  for (const [label, fn] of tasks) {
    try {
      fn();
      results.push(`${label}: OK`);
    } catch (err) {
      results.push(`${label}: ${err && err.message ? err.message : 'Error'}`);
    }
  }

  SpreadsheetApp.flush();
  try { ss.toast('Submit All finished:\n' + results.join('\n'), 'Submit All', 10); } catch (_){}
  try { lock.releaseLock(); } catch (_){}
}

function fillTodayDateIfBlank_(sheet) {
  const today = new Date();
  const dateOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate()); // strips time

  const ranges = [
    'B4',
    'E4',
    'I4:I8',
    'AF4',
    'AI4',
    'AS4',
    'AY4'
  ];

  ranges.forEach(a1 => {
    const r = sheet.getRange(a1);
    const rows = r.getNumRows();
    const cols = r.getNumColumns();
    const vals = r.getValues();
    let changed = false;

    for (let i = 0; i < rows; i++) {
      for (let j = 0; j < cols; j++) {
        if (vals[i][j] === '' || vals[i][j] === null) {
          vals[i][j] = dateOnly;
          changed = true;
        }
      }
    }

    if (changed) {
      r.setValues(vals);
      r.setNumberFormat('M/d/yyyy'); // Ensures no time shown
    }
  });
}

function ensureInputsDatesPresent_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Inputs');
  fillTodayDateIfBlank_(sh);
}

function runAutoDateFill() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Inputs');
  if (!sh) return;

  // 1) Clear the date input cells first
  sh.getRangeList([
    'B4',
    'E4',
    'I4:I8',
    'AF4',
    'AI4',
    'AS4',
    'AY4'
  ]).clearContent();

  // 2) Refill date cells if blank (will write today's date)
  fillTodayDateIfBlank_(sh);

  // 3) Clear recent input/daily total previews
  sh.getRangeList([
    'I17',
    'N17:AD17',
    'AF8:AG8',
    'AI8:AQ8',
    'AS8:AW8',
    'AY8:AZ8'
  ]).clearContent();

  // (Optional) keep any formats you want after clearing, e.g.:
  // sh.getRange('AF8').setNumberFormat('M/d/yyyy H:mm');
}

function submitWeight() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Inputs');
  const dataTableSheet = ss.getSheetByName('Data Tables');

  const dateTime = inputSheet.getRange('B4').getValue();
  const weight = inputSheet.getRange('C4').getValue();

  if (dateTime && weight) {
    // Find next empty row in Data Tables col B (starting at row 4)
    const values = dataTableSheet.getRange('B4:B' + dataTableSheet.getMaxRows()).getValues();
    let emptyRowOffset = values.findIndex(row => !row[0]);
    if (emptyRowOffset === -1) emptyRowOffset = values.length;
    const targetRow = 4 + emptyRowOffset;

    dataTableSheet.getRange(targetRow, 2).setValue(dateTime); // Col B
    dataTableSheet.getRange(targetRow, 3).setValue(weight);   // Col C

    // Show the latest submission in the preview row
    inputSheet.getRange('B8:C8').setValues([[dateTime, weight]]);

    // Clear input
    inputSheet.getRange('B4:C4').clearContent();

    // Refill date cells if blank (because onEdit won't fire on script changes)
    ensureInputsDatesPresent_();
  }
}

function submitBodyFat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Inputs');
  const dataTableSheet = ss.getSheetByName('Data Tables');

  // Get input values
  const val1 = inputSheet.getRange('E4').getValue(); // e.g., Date/Time or BF%
  const val2 = inputSheet.getRange('F4').getValue(); // e.g., something else
  const val3 = inputSheet.getRange('G4').getValue(); // e.g., something else

  // Only run if all fields have a value
  if (val1 && val2 && val3) {
    // Find first empty rows in F, G, I (from row 4)
    const colF = dataTableSheet.getRange('F4:F' + dataTableSheet.getMaxRows()).getValues();
    const colG = dataTableSheet.getRange('G4:G' + dataTableSheet.getMaxRows()).getValues();
    const colI = dataTableSheet.getRange('I4:I' + dataTableSheet.getMaxRows()).getValues();

    let rowF = colF.findIndex(row => !row[0]);
    let rowG = colG.findIndex(row => !row[0]);
    let rowI = colI.findIndex(row => !row[0]);

    if (rowF === -1) rowF = colF.length;
    if (rowG === -1) rowG = colG.length;
    if (rowI === -1) rowI = colI.length;

    // Keep the trio aligned: use the lowest available row index
    const targetRow = 4 + Math.max(rowF, rowG, rowI);

    dataTableSheet.getRange(targetRow, 6).setValue(val1); // Col F
    dataTableSheet.getRange(targetRow, 7).setValue(val2); // Col G
    dataTableSheet.getRange(targetRow, 9).setValue(val3); // Col I

    // Show the latest submission in the preview row
    inputSheet.getRange('E8:G8').setValues([[val1, val2, val3]]);

    // Clear inputs
    inputSheet.getRange('E4:G4').clearContent();

    // Refill date cells if blank (because onEdit won't fire on script changes)
    ensureInputsDatesPresent_();
  }
}

function prependTimeTablesBD_(sheet, dateTime, name, servings) {
  const START_ROW = 4;
  const startCol = 2; // B
  const width = 3;    // B:D

  // Get existing block
  const height = Math.max(0, sheet.getLastRow() - (START_ROW - 1));
  const existing = (height > 0)
    ? sheet.getRange(START_ROW, startCol, height, width).getValues()
    : [];

  // Filter out completely blank rows
  const compact = existing.filter(r => r.some(v => v !== '' && v !== null));

  // Normalize date to Date object
  const dt = (dateTime instanceof Date) ? dateTime : new Date(dateTime);

  // Build new row
  const newRow = [[dt, name, servings]];

  // Merge + sort by date desc
  const merged = compact.concat(newRow).sort((a, b) => new Date(b[0]) - new Date(a[0]));

  // Write back
  sheet.getRange(START_ROW, startCol, merged.length, width).setValues(merged);

  // Clear any leftover rows
  const leftover = height - merged.length;
  if (leftover > 0) {
    sheet.getRange(START_ROW + merged.length, startCol, leftover, width).clearContent();
  }
}

function submitNutrition() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Inputs');
  const dataTableSheet = ss.getSheetByName('Data Tables');
  const timeTableSheet = ss.getSheetByName('Time Tables');
  const foodRefSheet = ss.getSheetByName('Food Nutrition Reference');

  const inputs = inputSheet.getRange('I4:AD13').getValues(); // 10 rows x 23 cols
  let lastProcessedDateOnly = null; // remember the last date we touched so we can mirror totals to I17:AD17

  for (let i = 0; i < inputs.length; i++) {
    const row = inputs[i];
    const dateTime = row[0];
    const name = row[1];
    const type = row[2];
    const weightVolume = row[3];
    const servings = Number(row[4]) || 1;

    if (!dateTime || !name) continue; // skip empties

    // ====== 1) Log to Time Tables (B:D, newest on top, by date desc) ======
    prependTimeTablesBD_(timeTableSheet, dateTime, name, servings);

    // ====== 2) Build nutrition vector (17 cols) ======
    if (!type) continue; // only proceed if Food/Drink indicated
    const nameNormalized = name.toString().trim().toLowerCase();

    let nutritionData = null; // 17 numbers
    let isReference = false;

    if (type === 'Food') {
      const foodNames = foodRefSheet.getRange('B3:B1000').getValues().flat();
      const idx = foodNames.findIndex(n => (n ? n.toString().trim().toLowerCase() : '') === nameNormalized);
      if (idx !== -1) {
        const ref = foodRefSheet.getRange(3 + idx, 4, 1, 17).getValues()[0]; // D:T (17)
        nutritionData = ref.map(n => (Number(n) || 0) * servings);
        isReference = true;
      }
    } else if (type === 'Drink') {
      const drinkNames = foodRefSheet.getRange('V3:V1000').getValues().flat();
      const idx = drinkNames.findIndex(n => (n ? n.toString().trim().toLowerCase() : '') === nameNormalized);
      if (idx !== -1) {
        const ref = foodRefSheet.getRange(3 + idx, 24, 1, 17).getValues()[0]; // X:AN (17)
        nutritionData = ref.map(n => (Number(n) || 0) * servings);
        isReference = true;
      }
    }

    // If not found in reference, use Inputs N:AD (17 cols raw) and add to reference table
    if (!isReference) {
      const raw = row.slice(5, 22); // N:AD → 17 columns
      nutritionData = raw.map(n => (Number(n) || 0) * servings);

      if (type === 'Food') {
        const list = foodRefSheet.getRange('B3:B1000').getValues().flat();
        let emptyRow = list.findIndex(n => !n) + 3;
        if (emptyRow === 2) emptyRow = list.length + 3; // none empty, append
        foodRefSheet.getRange(emptyRow, 2).setValue(name);         // B (name)
        foodRefSheet.getRange(emptyRow, 3).setValue(weightVolume); // C (weight/volume)
        foodRefSheet.getRange(emptyRow, 4, 1, 17).setValues([raw]); // D:T (store raw, not multiplied)

        // Sort reference tables & refresh dropdowns
        sortReferenceTables_();
        const helper = ss.getSheetByName('_AutocompleteHelper');
        if (helper) writeHelperBlocks_(helper, getLiveRefNames_());

      } else if (type === 'Drink') {
        const list = foodRefSheet.getRange('V3:V1000').getValues().flat();
        let emptyRow = list.findIndex(n => !n) + 3;
        if (emptyRow === 2) emptyRow = list.length + 3;
        foodRefSheet.getRange(emptyRow, 22).setValue(name);        // V
        foodRefSheet.getRange(emptyRow, 23).setValue(weightVolume); // W
        foodRefSheet.getRange(emptyRow, 24, 1, 17).setValues([raw]); // X:AN

        // Sort reference tables & refresh dropdowns
        sortReferenceTables_();
        const helper = ss.getSheetByName('_AutocompleteHelper');
        if (helper) writeHelperBlocks_(helper, getLiveRefNames_());
      }
    }

    // ====== 3) Add/Sum into Data Tables for that date ======
    const dt = new Date(dateTime);
    const dateOnly = new Date(dt); dateOnly.setHours(0,0,0,0);
    lastProcessedDateOnly = dateOnly; // remember latest processed date

    const dateCol = dataTableSheet.getRange('L4:L1000').getValues().flat();
    let dateIdx = dateCol.findIndex(d => {
      if (!d) return false;
      const sd = new Date(d); sd.setHours(0,0,0,0);
      return sd.getTime() === dateOnly.getTime();
    });

    let targetRow;
    if (dateIdx !== -1) {
      targetRow = 4 + dateIdx;
      const existing = dataTableSheet.getRange(targetRow, 13, 1, 17).getValues()[0]; // M:AC
      const summed = existing.map((v, j) => (Number(v) || 0) + (Number(nutritionData[j]) || 0));
      dataTableSheet.getRange(targetRow, 13, 1, 17).setValues([summed]);
    } else {
      // new date row
      let emptyIdx = dateCol.findIndex(d => !d);
      if (emptyIdx === -1) emptyIdx = dateCol.length;
      targetRow = 4 + emptyIdx;
      dataTableSheet.getRange(targetRow, 12).setValue(dateOnly);               // L (date)
      dataTableSheet.getRange(targetRow, 13, 1, 17).setValues([nutritionData]); // M:AC (17)
    }
  }

  // ====== 4) Mirror the most recent processed date totals to Inputs!I17:AD17 ======
  if (lastProcessedDateOnly) {
    const dateCol = dataTableSheet.getRange('L4:L1000').getValues().flat();
    let dateIdx = dateCol.findIndex(d => {
      if (!d) return false;
      const sd = new Date(d); sd.setHours(0,0,0,0);
      return sd.getTime() === lastProcessedDateOnly.getTime();
    });
    if (dateIdx !== -1) {
      const row = 4 + dateIdx;
      const totals = dataTableSheet.getRange(row, 13, 1, 17).getValues()[0]; // M:AC
      inputSheet.getRange('I17').setValue(lastProcessedDateOnly);             // date
      inputSheet.getRange('N17:AD17').setValues([totals]);                    // 17 totals
    } else {
      // no row found (unlikely); clear preview
      inputSheet.getRange('I17').clearContent();
      inputSheet.getRange('N17:AD17').clearContent();
    }
  }

  // ====== 5) Clear input rows ======
  inputSheet.getRange('I4:AD13').clearContent();

  // Refill date cells if blank (because onEdit won't fire on script changes)
  ensureInputsDatesPresent_();
}

function sortReferenceTables_() {
  const ss  = SpreadsheetApp.getActive();
  const ref = ss.getSheetByName('Food Nutrition Reference');
  if (!ref) return;

  // Sort Food block (B:T) and Drink block (V:AN) by their name column (first col of each block)
  _safeSortBlockByFirstColumn(ref, 3,  2, 19);  // startRow=3, startCol=B(2), width up to T
  _safeSortBlockByFirstColumn(ref, 3, 22, 19);  // startRow=3, startCol=V(22), width up to AN
}

function _safeSortBlockByFirstColumn(sheet, startRow, startCol, maxWidth) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < startRow || lastCol < startCol) return;

  const endCol = Math.min(startCol + maxWidth - 1, lastCol);
  const width  = endCol - startCol + 1;
  const height = lastRow - startRow + 1;
  if (width <= 0 || height <= 0) return;

  const rng = sheet.getRange(startRow, startCol, height, width);

  // Kill merges in the block (merges + sorting = pain)
  const merges = rng.getMergedRanges();
  if (merges && merges.length) merges.forEach(m => m.breakApart());

  const values = rng.getValues();

  // Keep only rows with a name in the FIRST column of the block (B or V)
  const rowsWithNames = [];
  for (let i = 0; i < values.length; i++) {
    const name = values[i][0];
    if (name !== '' && name !== null) rowsWithNames.push(values[i]);
  }
  if (rowsWithNames.length === 0) return;

  // Sort by case-insensitive name
  rowsWithNames.sort((a, b) => {
    const A = a[0].toString().trim().toLowerCase();
    const B = b[0].toString().trim().toLowerCase();
    return A < B ? -1 : A > B ? 1 : 0;
  });

  // Clear the whole block, then write compacted, sorted rows back to the top
  rng.clearContent();
  sheet.getRange(startRow, startCol, rowsWithNames.length, width).setValues(rowsWithNames);
}

function trimColumnIndex_(sheet, colIndex, startRow) {
  const lastRow = sheet.getLastRow();
  const height  = Math.max(0, lastRow - startRow + 1);
  if (height <= 0 || colIndex > sheet.getMaxColumns()) return;
  const rng  = sheet.getRange(startRow, colIndex, height, 1);
  const vals = rng.getValues();
  let changed = false;
  for (let i = 0; i < vals.length; i++) {
    if (typeof vals[i][0] === 'string') {
      const t = vals[i][0].trim();
      if (t !== vals[i][0]) { vals[i][0] = t; changed = true; }
    }
  }
  if (changed) rng.setValues(vals);
}

function getLastNonEmptyRow_(sheet, colIndex, startRow) {
  const lastRow = sheet.getLastRow();
  const height  = Math.max(0, lastRow - startRow + 1);
  if (height <= 0 || colIndex > sheet.getMaxColumns()) return startRow - 1;
  const values = sheet.getRange(startRow, colIndex, height, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const v = values[i][0];
    if (v !== '' && v !== null) return startRow + i;
  }
  return startRow - 1;
}

/********** ONE-TIME SETUP **********/
function setupNutritionAutocomplete() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputs = ss.getSheetByName('Inputs');

  // Make (or clear) helper sheet
  let helper = ss.getSheetByName('_AutocompleteHelper');
  if (!helper) {
    helper = ss.insertSheet('_AutocompleteHelper');
    helper.hideSheet();
  } else {
    helper.clear();
  }

  // Build helper blocks with the full live list so dropdowns exist immediately
  const allNames = getLiveRefNames_();
  writeHelperBlocks_(helper, allNames);

  // Set permanent validation on Inputs!J4..J13 pointing at each block
  const BLOCK_SIZE = 1000;     // must match writeHelperBlocks_
  const BLOCK_SPACING = 1200;  // must match writeHelperBlocks_
  for (let i = 0; i < 10; i++) {
    const startRow = i * BLOCK_SPACING + 1;
    const range = helper.getRange(startRow, 1, BLOCK_SIZE, 1); // column A block
    const targetRow = 4 + i;
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(range, true) // allow other entries too
      .setHelpText('Suggestions from Food Nutrition Reference (fuzzy).')
      .build();
    inputs.getRange(targetRow, 10).setDataValidation(rule); // J column
  }
}

/********** WRITE helper blocks with the FULL list **********/
function writeHelperBlocks_(helper, allNames) {
  // Sort alphabetically for predictable dropdown when no typing
  const sorted = allNames.slice().sort((a, b) => a.localeCompare(b));

  const BLOCK_SIZE = 1000;     // up to 1000 suggestions per row
  const BLOCK_SPACING = 1200;  // spacing between blocks
  const base = sorted.slice(0, BLOCK_SIZE).map(v => [v]);
  while (base.length < BLOCK_SIZE) base.push(['']);

  // J4..J13 -> 10 blocks
  for (let i = 0; i < 10; i++) {
    const startRow = i * BLOCK_SPACING + 1;
    helper.getRange(startRow, 1, BLOCK_SIZE, 1).setValues(base);
  }
}

/********** LIVE lookup (no cache) **********/
function getLiveRefNames_() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ref = ss.getSheetByName('Food Nutrition Reference');
  if (!ref) return [];

  const lastRow = ref.getLastRow();
  const maxCols = ref.getMaxColumns();
  const height  = Math.max(0, lastRow - 3 + 1);

  const foods = (height > 0 && maxCols >= 2)
    ? ref.getRange(3, 2,  height, 1).getValues().flat().filter(Boolean)
    : [];

  const drinks = (height > 0 && maxCols >= 22)
    ? ref.getRange(3, 22, height, 1).getValues().flat().filter(Boolean)
    : [];

  return Array.from(new Set(
    foods.concat(drinks).map(n => n.toString().trim()).filter(Boolean)
  ));
}

/********** Update a single row's helper block while typing **********/
function updateNutritionSuggestionsBlock_(rowNumber, typed) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const helper = ss.getSheetByName('_AutocompleteHelper');
  if (!helper) return;

  const allNames = getLiveRefNames_();
  const q = normalize_(typed || '');
  if (!q) return; // leave the full list from writeHelperBlocks_ when no query

  // Fuzzy score: startsWith > contains > similarity
  const scored = allNames.map(name => {
    const ln = normalize_(name);
    let score = 0;
    if (ln.startsWith(q)) score += 3;
    if (ln.includes(q))   score += 2;
    score += similarity_(q, ln);
    return { name, score };
  });
  scored.sort((a, b) => b.score - a.score || a.name.localeCompare(b.name));

  const suggestions = scored.slice(0, 300).map(x => x.name); // top N while typing

  // Write suggestions to just that row's block
  const BLOCK_SIZE = 1000;
  const BLOCK_SPACING = 1200;
  const startRow = (rowNumber - 4) * BLOCK_SPACING + 1;
  const range = helper.getRange(startRow, 1, BLOCK_SIZE, 1);
  const out = suggestions.map(v => [v]);
  while (out.length < BLOCK_SIZE) out.push(['']);
  range.setValues(out);
}

/********** Fuzzy helpers **********/
function normalize_(s) {
  return s
    .toString()
    .toLowerCase()
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9 ]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function similarity_(a, b) {
  if (!a && !b) return 1;
  if (!a || !b) return 0;
  const dist = levenshtein_(a, b);
  const maxLen = Math.max(a.length, b.length);
  return maxLen ? (1 - dist / maxLen) : 0;
}

function levenshtein_(a, b) {
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;

  const prev = new Array(n + 1);
  const curr = new Array(n + 1);
  for (let j = 0; j <= n; j++) prev[j] = j;

  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    const ca = a.charCodeAt(i - 1);
    for (let j = 1; j <= n; j++) {
      const cost = (ca === b.charCodeAt(j - 1)) ? 0 : 1;
      curr[j] = Math.min(
        curr[j - 1] + 1,
        prev[j] + 1,
        prev[j - 1] + cost
      );
    }
    for (let j = 0; j <= n; j++) prev[j] = curr[j];
  }
  return prev[n];
}

/********** Refill date cells after script-driven clears **********/
function ensureInputsDatesPresent_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Inputs');
  fillTodayDateIfBlank_(sh);
}

function fillTodayDateIfBlank_(sheet) {
  const today = new Date();
  const dateOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate()); // strips time

  const ranges = ['B4', 'E4', 'I4:I8', 'AF4', 'AI4', 'AS4', 'AY4'];

  ranges.forEach(a1 => {
    const r = sheet.getRange(a1);
    const rows = r.getNumRows();
    const cols = r.getNumColumns();
    const vals = r.getValues();
    let changed = false;

    for (let i = 0; i < rows; i++) {
      for (let j = 0; j < cols; j++) {
        if (vals[i][j] === '' || vals[i][j] === null) {
          vals[i][j] = dateOnly;
          changed = true;
        }
      }
    }

    if (changed) {
      r.setValues(vals);
      r.setNumberFormat('M/d/yyyy'); // Ensures no time shown
    }
  });
}

function mirrorNutritionPreviewFromRow_(row) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const dt  = ss.getSheetByName('Data Tables');
  const inp = ss.getSheetByName('Inputs');
  if (!dt || !inp) return;

  const dateVal = dt.getRange(row, 12).getValue(); // L
  if (!dateVal) {
    inp.getRange('I17').clearContent();
    inp.getRange('N17:AD17').clearContent();
    return;
  }
  const totals = dt.getRange(row, 13, 1, 17).getValues()[0]; // M:AC
  const dateOnly = new Date(dateVal); dateOnly.setHours(0,0,0,0);

  inp.getRange('I17').setValue(dateOnly);
  inp.getRange('N17:AD17').setValues([totals]);
}

function submitWater() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Inputs');
  const dataTableSheet = ss.getSheetByName('Data Tables');

  const [inputDate, inputWater] = inputSheet.getRange('AF4:AG4').getValues()[0];
  if (!inputDate || !inputWater) return;

  const inputDateObj = new Date(inputDate);
  inputDateObj.setHours(0,0,0,0);

  const data = dataTableSheet.getRange(4, 31, Math.max(0, dataTableSheet.getLastRow() - 3), 2).getValues();
  let matchedRow = null;
  let firstBlankRow = null;

  for (let i = 0; i < data.length; i++) {
    const [rowDate, rowWater] = data[i];
    if (firstBlankRow === null && !rowDate && !rowWater) firstBlankRow = i + 4;
    if (rowDate) {
      const rowDateObj = new Date(rowDate);
      rowDateObj.setHours(0,0,0,0);
      if (rowDateObj.getTime() === inputDateObj.getTime()) {
        matchedRow = i + 4;
        break;
      }
    }
  }

  let totalForDay, dateForDisplay;

  if (matchedRow !== null) {
    const existingWater = Number(dataTableSheet.getRange(matchedRow, 32).getValue()) || 0; // col AF? actually col 32 = AG
    totalForDay = existingWater + Number(inputWater);
    dataTableSheet.getRange(matchedRow, 32).setValue(totalForDay);
    dateForDisplay = dataTableSheet.getRange(matchedRow, 31).getValue(); // keep the stored date
  } else {
    const appendRow = firstBlankRow ? firstBlankRow : dataTableSheet.getLastRow() + 1;
    totalForDay = Number(inputWater);
    dataTableSheet.getRange(appendRow, 31, 1, 2).setValues([[inputDate, totalForDay]]);
    dateForDisplay = inputDateObj; // date we just inserted
  }

  const triggerTime = new Date();
  const af8 = new Date(dateForDisplay); // the day you’re logging water for
  af8.setHours(triggerTime.getHours(), triggerTime.getMinutes(), triggerTime.getSeconds(), 0);

  inputSheet.getRange('AF8').setValue(af8);
  inputSheet.getRange('AG8').setValue(totalForDay);
  inputSheet.getRange('AF8').setNumberFormat("M/d/yyyy H:mm");

  // Clear entry inputs
  inputSheet.getRange('AF4:AG4').clearContent();

  // Refill date cells if blank (because onEdit won't fire on script changes)
    ensureInputsDatesPresent_();
}

function parseYesCount(cellVal) {
  if (!cellVal) return 0;
  const str = cellVal.toString().trim().toLowerCase();
  if (str === "yes") return 1;
  const match = str.match(/^yes\s*x\s*(\d+)$/);
  if (match) return parseInt(match[1], 10);
  return 0;
}

// NEW: parse a user entry like "Yes", "Y", "Yes x2", "y x3", or even "2" into an integer count
function parseYesInputCount(val) {
  if (!val) return 0;
  const s = val.toString().trim().toLowerCase();
  if (s === "yes" || s === "y") return 1;
  let m = s.match(/^y(?:es)?\s*x\s*(\d+)$/);  // "yes x2" or "y x3"
  if (m) return Math.max(1, parseInt(m[1], 10));

  // plain number like "2" (treat as Yes x2)
  if (/^\d+$/.test(s)) return Math.max(0, parseInt(s, 10));

  // fallback: 0 (includes "n/a" which caller handles)
  return 0;
}

// NEW: multiply a dose by a count and keep units if present
function multiplyDose(doseVal, count) {
  if (!doseVal || count <= 0) return "";
  // If numeric cell
  if (typeof doseVal === 'number') {
    return doseVal * count;
  }
  // If string like "5000 mg" or "5 g"
  const s = doseVal.toString().trim();
  const num = parseFloat(s.replace(/,/g, ''));
  if (isNaN(num)) {
    // unknown format -> best effort: "dose xN"
    return count === 1 ? s : (s + " x" + count);
  }
  // Extract unit (letters and %), e.g. "mg", "g", "%". Remove spaces between number & unit.
  const unitMatch = s.replace(/,/g, '').match(/[a-zA-Z%]+/g);
  const unit = unitMatch ? unitMatch.join('') : '';
  const total = num * count;
  // Avoid trailing .0
  const totalStr = Number.isInteger(total) ? String(total) : String(total);
  return unit ? (totalStr + unit) : total;
}

// Prepend + keep newest-first for Time Tables F:H (Supplements).
// rows = array of [dateTime, name, dose]
function prependTimeTablesFH_(sheet, rows) {
  if (!rows || rows.length === 0) return;

  const START_ROW = 4;
  const startCol  = 6; // F
  const width     = 3; // F:H

  const height   = Math.max(0, sheet.getLastRow() - (START_ROW - 1));
  const existing = (height > 0)
    ? sheet.getRange(START_ROW, startCol, height, width).getValues()
    : [];

  const compact = existing.filter(r => r.some(v => v !== '' && v !== null));

  const normalizedAdds = rows.map(r => {
    const dt = (r[0] instanceof Date) ? r[0] : new Date(r[0]);
    const out = r.slice(); out[0] = dt; return out;
  });

  const merged = compact.concat(normalizedAdds)
    .sort((a, b) => new Date(b[0]) - new Date(a[0]));

  sheet.getRange(START_ROW, startCol, merged.length, width).setValues(merged);

  const leftover = height - merged.length;
  if (leftover > 0) {
    sheet.getRange(START_ROW + merged.length, startCol, leftover, width).clearContent();
  }
}

function submitSupplements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Inputs');
  const dataTableSheet = ss.getSheetByName('Data Tables');
  const timeTableSheet = ss.getSheetByName('Time Tables');
  const supplementsSheet = ss.getSheetByName('Supplements');

  // --- Inputs ---
  const [inputDate, ...inputs] = inputSheet.getRange('AI4:AQ4').getValues()[0];
  if (!inputDate) return;
  const names = inputSheet.getRange('AJ3:AQ3').getValues()[0];

  // --- Data Tables: find date row for Supplements log (AH:AP) ---
  const logHeight = Math.max(0, dataTableSheet.getLastRow() - 3);
  const log = logHeight > 0
    ? dataTableSheet.getRange(4, 34, logHeight, 9).getValues()  // AH4:AP
    : [];

  const inputDateObj = new Date(inputDate); inputDateObj.setHours(0,0,0,0);
  let matchedRow = null;
  let firstBlankRow = null;
  for (let i = 0; i < log.length; i++) {
    const rowDate = log[i][0];
    if (firstBlankRow === null && !rowDate) firstBlankRow = i + 4;
    if (rowDate) {
      const d = new Date(rowDate); d.setHours(0,0,0,0);
      if (d.getTime() === inputDateObj.getTime()) {
        matchedRow = i + 4;
        break;
      }
    }
  }

  // --- Update tally in Data Tables with Yes/Yes xN/No/N/A logic ---
  if (matchedRow !== null) {
    for (let j = 0; j < 8; j++) {
      const vRaw = inputs[j];
      const v = vRaw ? vRaw.toString().trim().toLowerCase() : "";
      const cell = dataTableSheet.getRange(matchedRow, 35 + j); // AI..AP

      if (v === "n/a") { cell.setValue("N/A"); continue; }
      if (v === "no" || v === "n") { cell.setValue("No"); continue; }

      const addCount = parseYesInputCount(vRaw);
      if (addCount > 0) {
        const existingVal = cell.getValue();
        const existingCount = parseYesCount(existingVal);
        const newCount = existingCount + addCount;
        cell.setValue(newCount === 1 ? "Yes" : ("Yes x" + newCount));
      }
    }
  } else if (firstBlankRow) {
    // Build initial row for new date: AH=date, AI:AP per inputs (Yes/Yes xN/No/N/A/"")
    const out = [inputDate];
    for (let j = 0; j < 8; j++) {
      const vRaw = inputs[j];
      const v = vRaw ? vRaw.toString().trim().toLowerCase() : "";
      if (v === "n/a") {
        out.push("N/A");
      } else if (v === "no" || v === "n") {
        out.push("No");
      } else {
        const addCount = parseYesInputCount(vRaw);
        if (addCount <= 0) out.push("");
        else out.push(addCount === 1 ? "Yes" : ("Yes x" + addCount));
      }
    }
    dataTableSheet.getRange(firstBlankRow, 34, 1, 9).setValues([out]);
    matchedRow = firstBlankRow;
  }

  // --- Build supplement dose lookup ---
  const suppVals = supplementsSheet.getRange('B4:C').getValues();
  const doseByName = {};
  for (let i = 0; i < suppVals.length; i++) {
    const n = suppVals[i][0], d = suppVals[i][1];
    if (!n) break;
    doseByName[n.toString().trim().toLowerCase()] = d;
  }

  // --- Collect YES events to prepend to Time Tables, and count Vitamin D3 ---
  // F = DateTime, G = Name, H = Dose (multiplied by "Yes xN" count)
  const newLogRows = [];
  let d3YesCount = 0; // add 1.0 per "Yes" (handles "Yes xN")

  for (let j = 0; j < 8; j++) {
    const raw = inputs[j];
    const s = raw ? raw.toString().trim().toLowerCase() : "";

    if (s === "n/a" || s === "no" || s === "n") continue; // skip logging

    const addCount = parseYesInputCount(raw);
    if (addCount > 0) {
      const nm = names[j] || "";
      const nmStr = nm.toString();
      const baseDose = doseByName[nmStr.trim().toLowerCase()] ?? "";
      const totalDose = multiplyDose(baseDose, addCount);
      newLogRows.push([inputDate, nmStr, totalDose]);

      if (/\bvitamin\s*d3\b/i.test(nmStr)) d3YesCount += addCount;
    }
  }

  // Write supplements time log (F:H) newest-first, in-memory
  if (newLogRows.length > 0) {
    prependTimeTablesFH_(timeTableSheet, newLogRows);
  }

  // --- Add +1.0 per Vitamin D3 "Yes" to Nutrition Vitamin D (Data Tables!AA) ---
  if (d3YesCount > 0) {
    const nutrHeight = Math.max(0, dataTableSheet.getLastRow() - 3);
    const datesCol = nutrHeight > 0
      ? dataTableSheet.getRange(4, 12, nutrHeight, 1).getValues().flat() // L
      : [];

    let dateIdx = -1, firstBlank = -1;
    for (let i = 0; i < datesCol.length; i++) {
      const v = datesCol[i];
      if (!v) { if (firstBlank === -1) firstBlank = i; continue; }
      const d = new Date(v); d.setHours(0,0,0,0);
      if (d.getTime() === inputDateObj.getTime()) { dateIdx = i; break; }
    }

    const targetRow = (dateIdx !== -1)
      ? 4 + dateIdx
      : 4 + (firstBlank !== -1 ? firstBlank : datesCol.length);

    if (dateIdx === -1) {
      dataTableSheet.getRange(targetRow, 12).setValue(inputDateObj); // L = date-only
    }

    const vitDCell = dataTableSheet.getRange(targetRow, 27); // AA
    const existing = Number(vitDCell.getValue()) || 0;
    vitDCell.setValue(existing + d3YesCount * 1); // 1.0 per "Yes"
    // vitDCell.setNumberFormat("0%"); // optional
    mirrorNutritionPreviewFromRow_(targetRow);
  }

  // --- Mirror current date's supplements row to Inputs!AI8:AQ8 ---
  let previewRow = matchedRow;
  if (!previewRow) {
    const ahDatesHeight = Math.max(0, dataTableSheet.getLastRow() - 3);
    const dates = ahDatesHeight > 0
      ? dataTableSheet.getRange(4, 34, ahDatesHeight, 1).getValues().flat() // AH
      : [];
    for (let i = 0; i < dates.length; i++) {
      const d = dates[i]; if (!d) continue;
      const dd = new Date(d); dd.setHours(0,0,0,0);
      if (dd.getTime() === inputDateObj.getTime()) { previewRow = i + 4; break; }
    }
  }
  if (previewRow) {
    const rowVals = dataTableSheet.getRange(previewRow, 34, 1, 9).getValues(); // AH:AP
    inputSheet.getRange('AI8:AQ8').setValues(rowVals);
  } else {
    inputSheet.getRange('AI8:AQ8').clearContent();
  }

  // Clear Inputs & ensure dates
  inputSheet.getRange('AI4:AQ4').clearContent();
  ensureInputsDatesPresent_();
}

// Prepend + keep newest-first for Time Tables J:K (Skincare).
// rows = array of [dateTime, productName]
function prependTimeTablesJK_(sheet, rows) {
  if (!rows || rows.length === 0) return;

  const START_ROW = 4;
  const startCol  = 10; // J
  const width     = 2;  // J:K

  const height   = Math.max(0, sheet.getLastRow() - (START_ROW - 1));
  const existing = (height > 0)
    ? sheet.getRange(START_ROW, startCol, height, width).getValues()
    : [];

  const compact = existing.filter(r => r.some(v => v !== '' && v !== null));

  const normalizedAdds = rows.map(r => {
    const dt = (r[0] instanceof Date) ? r[0] : new Date(r[0]);
    const out = r.slice(); out[0] = dt; return out;
  });

  const merged = compact.concat(normalizedAdds)
    .sort((a, b) => new Date(b[0]) - new Date(a[0]));

  sheet.getRange(START_ROW, startCol, merged.length, width).setValues(merged);

  const leftover = height - merged.length;
  if (leftover > 0) {
    sheet.getRange(START_ROW + merged.length, startCol, leftover, width).clearContent();
  }
}

function submitSkincare() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Inputs');
  const dataTableSheet = ss.getSheetByName('Data Tables');
  const timeTableSheet = ss.getSheetByName('Time Tables');

  // Read input row: AS4:AW4 (DateTime + Toner, AM Moisturizer, Facial Cleanser, Oil Control)
  const [dateTime, toner, amMoist, cleanser, oilCtrl] = inputSheet.getRange('AS4:AW4').getValues()[0];
  if (!dateTime) return;

  // Headers for product names (AT3:AW3)
  const productNames  = inputSheet.getRange('AT3:AW3').getValues()[0];
  const productInputs = [toner, amMoist, cleanser, oilCtrl];

  // ---- 1) Data Tables: ensure/locate date row in AR (date only), then set/update AS:AV ----
  const dateOnly = new Date(dateTime);
  dateOnly.setHours(0,0,0,0);

  const arCol = dataTableSheet.getRange('AR4:AR1000').getValues().flat();
  let rowIdx = arCol.findIndex(v => {
    if (!v) return false;
    const d = new Date(v); d.setHours(0,0,0,0);
    return d.getTime() === dateOnly.getTime();
  });

  if (rowIdx === -1) {
    let blankIdx = arCol.findIndex(v => !v);
    if (blankIdx === -1) blankIdx = arCol.length;
    rowIdx = blankIdx;
    dataTableSheet.getRange(4 + rowIdx, 44).setValue(dateOnly); // AR=44
    dataTableSheet.getRange(4 + rowIdx, 45, 1, 4).clearContent(); // AS:AV
  }

  // Update each product column:
  // - "yes"/"y"  -> increment Yes/Yes xN tally
  // - "no"/"n"   -> write "No" (do not log to Time Tables)
  // - "n/a"      -> write "N/A" (do not log)
  for (let j = 0; j < 4; j++) {
    const valStr = (productInputs[j] || '').toString().trim().toLowerCase();
    if (!valStr) continue;

    const cell = dataTableSheet.getRange(4 + rowIdx, 45 + j); // AS..AV

    if (valStr === 'n/a') { cell.setValue('N/A'); continue; }
    if (valStr === 'no' || valStr === 'n') { cell.setValue('No'); continue; }
    if (valStr === 'yes' || valStr === 'y') {
      const existing = cell.getValue();
      const count = parseYesCount(existing) + 1;
      cell.setValue(count === 1 ? 'Yes' : ('Yes x' + count));
    }
  }

  // ---- 2) Time Tables: prepend logs to J:K (most recent on top) ONLY for YES/Y ----
  const newRows = [];
  for (let i = 0; i < productInputs.length; i++) {
    const valStr = (productInputs[i] || '').toString().trim().toLowerCase();
    if (valStr === 'yes' || valStr === 'y') {
      newRows.push([dateTime, productNames[i]]);
    }
  }
  if (newRows.length > 0) {
    prependTimeTablesJK_(timeTableSheet, newRows);
  }

  // ---- 3) Mirror current day's up-to-date data into Inputs!AS8:AW8 ----
  const snapshot = dataTableSheet.getRange(4 + rowIdx, 44, 1, 5).getValues()[0]; // AR..AV
  inputSheet.getRange('AS8:AW8').setValues([snapshot]); // AS=date, AT..AW tallies/statuses

  // Clear the skincare inputs & ensure dates
  inputSheet.getRange('AS4:AW4').clearContent();
  ensureInputsDatesPresent_();
}

function submitVyvanse() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Inputs');
  const dataTableSheet = ss.getSheetByName('Data Tables');

  // Read Inputs!AY4:AZ4 -> [dateTime, yesNo]
  const [dateTime, yesNo] = inputSheet.getRange('AY4:AZ4').getValues()[0];
  if (!dateTime) return; // need a timestamp to proceed

  // Prepare date-only and time-only
  const dt = new Date(dateTime);
  const dateOnly = new Date(dt); dateOnly.setHours(0,0,0,0);
  const timeOnly = new Date(dt); // will rely on cell format for time display

  // Find next available row in Data Tables AX:AX (start at row 4)
  const axCol = dataTableSheet.getRange('AX4:AX1000').getValues().flat();
  let blankIdx = axCol.findIndex(v => !v);
  if (blankIdx === -1) blankIdx = axCol.length; // append if full
  const targetRow = 4 + blankIdx;

  // Write AX (date only), AY (Yes/No), AZ (time only)
  dataTableSheet.getRange(targetRow, 50).setValue(dateOnly);      // AX
  dataTableSheet.getRange(targetRow, 51).setValue(yesNo || "");   // AY
  dataTableSheet.getRange(targetRow, 52).setValue(timeOnly);      // AZ

  // Apply formats
  dataTableSheet.getRange(targetRow, 50).setNumberFormat("M/d/yyyy"); // AX date
  dataTableSheet.getRange(targetRow, 52).setNumberFormat("H:mm");     // AZ time (24h)

  // === Mirror preview to Inputs!AY8:AZ8 ===
  // AY8 = full Date & Time; AZ8 = Yes/No
  inputSheet.getRange('AY8').setValue(dt);
  inputSheet.getRange('AZ8').setValue(yesNo || "");
  inputSheet.getRange('AY8').setNumberFormat("M/d/yyyy H:mm");
  inputSheet.getRange('AZ8').setNumberFormat("@");

  // Clear input cells
  inputSheet.getRange('AY4:AZ4').clearContent();

  // Refill date cells if blank (because onEdit won't fire on script changes)
    ensureInputsDatesPresent_();
}

function refreshTodayPreviewBlocks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inp = ss.getSheetByName('Inputs');
  const dt  = ss.getSheetByName('Data Tables');
  if (!inp || !dt) return;

  // Today (date-only)
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  today.setHours(0,0,0,0);

  // Helper: find row index (1-based) in Data Tables matching a date column
  function findRowByDate(colA1) {
    const rng = dt.getRange(colA1 + '4:' + colA1 + '1000').getValues().flat();
    for (let i = 0; i < rng.length; i++) {
      const v = rng[i];
      if (!v) continue;
      const d = new Date(v); d.setHours(0,0,0,0);
      if (d.getTime() === today.getTime()) return 4 + i;
    }
    return null;
  }

  // --- Nutrition preview (Inputs!I17 + N17:AD17) from Data Tables L(date), M:AC(totals)
  (function() {
    const row = findRowByDate('L');
    if (row) {
      const totals = dt.getRange(row, 13, 1, 17).getValues()[0]; // M:AC
      inp.getRange('I17').setValue(today);
      inp.getRange('I17').setNumberFormat('M/d/yyyy');
      inp.getRange('N17:AD17').setValues([totals]);
    } else {
      inp.getRange('I17').clearContent();
      inp.getRange('N17:AD17').clearContent();
    }
  })();

  // --- Water preview (Inputs!AF8:AG8) from Data Tables AE(date, col 31), AF(total, col 32)
  (function() {
    const lastRow = dt.getLastRow();
    const height  = Math.max(0, lastRow - 3);
    if (height === 0) { inp.getRange('AF8:AG8').clearContent(); return; }

    const water = dt.getRange(4, 31, height, 2).getValues(); // AE:AF
    let rowIdx = null;
    for (let i = 0; i < water.length; i++) {
      const v = water[i][0];
      if (!v) continue;
      const d = new Date(v); d.setHours(0,0,0,0);
      if (d.getTime() === today.getTime()) { rowIdx = i; break; }
    }
    if (rowIdx !== null) {
      const total = Number(water[rowIdx][1]) || 0; // col 32
      // Show today's date & current time for the preview
      const stamp = new Date();
      inp.getRange('AF8').setValue(stamp).setNumberFormat('M/d/yyyy H:mm');
      inp.getRange('AG8').setValue(total);
    } else {
      inp.getRange('AF8:AG8').clearContent();
    }
  })();

  // --- Supplements preview (Inputs!AI8:AQ8) from Data Tables AH:AP (34..42)
  (function() {
    const row = findRowByDate('AH');
    if (row) {
      const vals = dt.getRange(row, 34, 1, 9).getValues(); // AH:AP
      inp.getRange('AI8:AQ8').setValues(vals);
    } else {
      inp.getRange('AI8:AQ8').clearContent();
    }
  })();

  // --- Skincare preview (Inputs!AS8:AW8) from Data Tables AR:AV (44..48)
  (function() {
    const row = findRowByDate('AR');
    if (row) {
      const vals = dt.getRange(row, 44, 1, 5).getValues(); // AR:AV
      inp.getRange('AS8:AW8').setValues(vals);
    } else {
      inp.getRange('AS8:AW8').clearContent();
    }
  })();

  // --- Vyvanse preview (Inputs!AY8:AZ8)
  // Data Tables: AX(date-only, col 50), AY(Yes/No, col 51), AZ(time-only, col 52)
  (function() {
    const row = findRowByDate('AX');
    if (row) {
      const yesNo = dt.getRange(row, 51).getValue(); // AY
      const time  = dt.getRange(row, 52).getValue(); // AZ (time)
      // Combine today's date with stored time (if any)
      let when = new Date(today);
      if (time) {
        const t = new Date(time);
        when.setHours(t.getHours(), t.getMinutes(), t.getSeconds(), 0);
      }
      inp.getRange('AY8').setValue(when).setNumberFormat('M/d/yyyy H:mm');
      inp.getRange('AZ8').setValue(yesNo || '');
    } else {
      inp.getRange('AY8:AZ8').clearContent();
    }
  })();
}

function sortTimeTablesBlock_(sheet, startCol, numCols, ascending) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const height  = Math.max(0, lastRow - 3); // rows from 4..lastRow
  const width   = Math.min(numCols, Math.max(0, lastCol - startCol + 1));
  if (height <= 0 || width <= 0) return;

  const rng = sheet.getRange(4, startCol, height, width);
  const merges = rng.getMergedRanges();
  if (merges && merges.length) merges.forEach(m => m.breakApart());

  rng.sort([{ column: 1, ascending: !!ascending }]); // sort by leftmost col
}
