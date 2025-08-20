/************************  CONFIG  ************************/
const START_DATE = new Date('2025-01-01T00:00:00');
const SHEET_NAME = 'Cardio Data';
const KM_TO_MI   = 0.621371;
const MAX_DAYS   = 90;
const DATA_ROWS  = 31;               // max days in month
const BLOCK_ROWS = 34;               // 1 hdr + 1 titles + 31 data + 1 total
/*********************************************************/

/** Main entry */
function syncFit() {
  const sheet = ensureSheet_();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const buckets = fetchDailyBuckets_(START_DATE, today);
  writeGrid_(sheet, buckets);
}

/* ────────── SHEET UTIL ────────── */
function ensureSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

/* ────────── WRITE ENGINE ────────── */
function writeGrid_(sheet, buckets) {
  const tz          = Session.getScriptTimeZone();
  const monthNames  = ['January','February','March','April','May','June',
                       'July','August','September','October','November','December'];
  const byYM = {};  // year → month → { rows:[], sumSteps, sumKm }

  /* 1️⃣  Collect rows + running totals */
  buckets.forEach(b => {
    const d       = new Date(+b.startTimeMillis);
    const iso     = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    if (iso < Utilities.formatDate(START_DATE, tz, 'yyyy-MM-dd')) return;

    const year    = d.getFullYear();
    const month   = ('0'+(d.getMonth()+1)).slice(-2);             // '01'..'12'
    const mmddTxt = "'" + Utilities.formatDate(d, tz, 'MM/dd');

    const steps   = b.dataset[0].point[0]?.value[0]?.intVal || 0;
    const km      = (b.dataset[1].point[0]?.value[0]?.fpVal || 0)/1000;

    const ymObj = byYM[year] = byYM[year] || {};
    const mObj  = ymObj[month] = ymObj[month] || {rows:[], sumSteps:0, sumKm:0};

    mObj.rows.push([mmddTxt, steps, +(km*KM_TO_MI).toFixed(2), +km.toFixed(2)]);
    mObj.sumSteps += steps;
    mObj.sumKm    += km;
  });

  /* 2️⃣  Write each year block (34 rows) */
  Object.keys(byYM).sort().forEach((year, yIdx) => {
    const headerRow = 1 + yIdx * BLOCK_ROWS;      // 1, 35, 69, …
    const titleRow  = headerRow + 1;              // 2, 36, 70 …
    const dataStart = headerRow + 2;              // 3, 37, 71 …
    const totalRow  = headerRow + 2 + DATA_ROWS;  // 34, 68, 102 …

    /* 2a. Month headers (row 1) */
    monthNames.forEach((name, mIdx) => {
      const colStart = mIdx * 4 + 1;
      sheet.getRange(headerRow, colStart, 1, 4)
           .merge()
           .setValue(`${name} ${year}`)
           .setFontWeight('bold')
           .setHorizontalAlignment('center');
    });

    /* 2b. Column titles (row 2) */
    monthNames.forEach((_, mIdx) => {
      const colStart = mIdx * 4 + 1;
      sheet.getRange(titleRow, colStart, 1, 4)
           .setValues([['Date','Steps','Distance Mi','Distance Km']])
           .setFontWeight('bold');
    });

    /* 2c. Data + totals per month */
    for (let m = 1; m <= 12; m++) {
      const mKey     = ('0'+m).slice(-2);
      const info     = (byYM[year][mKey]) || {rows:[], sumSteps:0, sumKm:0};
      const rows     = info.rows.slice(0, DATA_ROWS);                 // max 31
      const pad      = DATA_ROWS - rows.length;
      if (pad > 0) rows.push(...Array.from({length:pad}, _=>['','','','']));

      const colStart = (m - 1)*4 + 1;

      /* write data rows (rows 3-33) */
      sheet.getRange(dataStart, colStart, DATA_ROWS, 4).setValues(rows);

      /* Bold the MM/dd column */
      sheet.getRange(dataStart, colStart, DATA_ROWS, 1).setFontWeight('bold');

      /* totals (row 34) */
      sheet.getRange(totalRow, colStart, 1, 4)
           .setValues([[
             'Total',
             info.sumSteps,
             +(info.sumKm * KM_TO_MI).toFixed(2),
             +info.sumKm.toFixed(2)
           ]])
           .setFontWeight('bold');
    }
  });
}

/* ────────── FIT FETCH (≤90-day chunks) ────────── */
function fetchDailyBuckets_(startIncl, endIncl) {
  const ONE_DAY = 864e5;
  const endExc  = new Date(endIncl.getTime() + ONE_DAY);
  const out = [];

  for (let cur = new Date(startIncl); cur < endExc; ) {
    const chunkEndExc = new Date(Math.min(
      cur.getTime() + MAX_DAYS*ONE_DAY, endExc.getTime()));

    const req = {
      aggregateBy:[
        {dataTypeName:'com.google.step_count.delta'},
        {dataTypeName:'com.google.distance.delta'}
      ],
      bucketByTime:{durationMillis:ONE_DAY},
      startTimeMillis:cur.getTime(),
      endTimeMillis:chunkEndExc.getTime()
    };

    const res = UrlFetchApp.fetch(
      'https://www.googleapis.com/fitness/v1/users/me/dataset:aggregate',{
        method:'post',
        contentType:'application/json',
        payload:JSON.stringify(req),
        headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()}
      });

    out.push(...(JSON.parse(res).bucket || []));
    cur = chunkEndExc;
  }
  return out;
}
