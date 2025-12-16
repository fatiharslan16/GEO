/**********************************************************
 * SyncRecitalUI (safe)
 * - Adds missing CSU IDs; fills student-info fields only
 * - Never overwrites schedule fields
 * - Sorts by Priority after insertions
 **********************************************************/
function syncRecitalUI() {
  const OVERWRITE = false; // set true if you want to refresh even non-empty cells

  const ss   = SpreadsheetApp.getActive();
  const ui   = (typeof UI_SHEET_NAME !== 'undefined'
                  ? ss.getSheetByName(UI_SHEET_NAME)
                  : ss.getSheetByName('RecitalUI')) || ensureUiSheet_();
  const resp = ss.getSheetByName(typeof RESPONSES_NAME !== 'undefined'
                  ? RESPONSES_NAME
                  : 'Student Recital Responses');
  if (!ui)   throw new Error('RecitalUI sheet not found.');
  if (!resp) throw new Error('Responses sheet not found.');

  // ---- headers & rows from Responses ----
  const rLc = resp.getLastColumn();
  const rLr = resp.getLastRow();
  if (rLr < 2) { ss.toast('No responses to sync.'); return; }
  const rHeaders = resp.getRange(1,1,1,rLc).getDisplayValues()[0].map(String);
  const rIdx = Object.fromEntries(rHeaders.map((h,i)=>[h,i]));
  const rRows = resp.getRange(2,1,rLr-1,rLc).getDisplayValues();

  // ---- headers from RecitalUI ----
  const uLc = ui.getLastColumn();
  const uHeaders = ui.getRange(1,1,1,uLc).getDisplayValues()[0].map(h=>String(h||'').trim());
  const uIdx = Object.fromEntries(uHeaders.map((h,i)=>[h,i])); // 0-based
  const lrUI0 = ui.getLastRow();

  // Fields to NEVER touch (assigned schedule)
  const PROTECT = new Set(['Date','Time','Access Time','Exit Time','Room','docFile']);

  // Map: Responses header -> RecitalUI header
  // Add any “one more field” here when you notice it’s missing.
  const COPY_MAP = [
    ['CSU ID',                    'CSU ID'],
    ['First Name',                'First Name'],
    ['Last Name',                 'Last Name'],
    ['Instrument or Voice',       'Instrument'],       // fallback handled below
    ['Applied Teacher Name',      'Applied Teacher'],  // fallback handled below
    ['Collaborative Pianist Name','Accompanist Name'], // fallback handled below
    ['Recital Length',            'Recital Length'],
    ['Reception',                 'Reception'],
    ['Phone Number',              'Phone Number'],
    ['Email Address',             'Email Address'],
    ['Recital Date',              'Recital Date'],     // student preference text/date
    ['Type of Recital',           'Recital Type']      // <- your new one
    // If you discover another field in Responses you want in UI:
    // ['Some Responses Header',    'Matching UI Header'],
  ];

  let created = 0, updated = 0;

  // Build easy getter that tries several names (for fallbacks)
  const getFromResp = (row, ...names) => {
    for (const n of names) {
      const i = rIdx[n];
      if (i != null) {
        const v = row[i];
        const s = (v == null ? '' : String(v).trim());
        if (s) return s;
      }
    }
    return '';
  };

  // Iterate responses IN SHEET ORDER → earlier rows = higher priority
  rRows.forEach(r => {
    const rawCsu = getFromResp(r, 'CSU ID'); if (!rawCsu) return;

    // Normalize to your display formats
    const csuFmt = formatCsuId_(rawCsu);
    const phone  = formatPhone_(getFromResp(r,'Phone Number','Phone'));
    const email  = getFromResp(r,'Email Address','Email');
    const first  = getFromResp(r,'First Name');
    const last   = getFromResp(r,'Last Name');
    const instr  = getFromResp(r,'Instrument or Voice','Instrument');
    const teach  = getFromResp(r,'Applied Teacher Name','Applied Teacher');
    const piano  = getFromResp(r,'Collaborative Pianist Name','Accompanist Name');
    const length = normalizeLength_(getFromResp(r,'Recital Length','Length'));
    const recept = getFromResp(r,'Reception','Reception Requested?','Reception Preference');
    const reqDt  = getFromResp(r,'Recital Date','Preferred Recital Date');
    const rtype  = getFromResp(r,'Type of Recital','Recital Type','Recital Category');

    // Ensure a row exists in RecitalUI for this CSU
    let row = findRowByCsu_(ui, rawCsu, csuFmt);
    if (!row) {
      row = Math.max(ui.getLastRow()+1, 2);
      // Set Priority (next available) and CSU ID
      ui.getRange(row, 1).setValue(nextPriority_(ui)); // Priority in col A
      const csuCol = (uIdx['CSU ID'] != null) ? (uIdx['CSU ID']+1) : 2; // default B
      ui.getRange(row, csuCol).setValue(csuFmt);
      created++;
    }

    // Read the whole UI row to update in-memory then write once
    const vals = ui.getRange(row, 1, 1, uLc).getDisplayValues()[0];

    // Helper: set if allowed and empty (unless OVERWRITE=true)
    const setIf = (uiHeader, value, always=false) => {
      if (!value && !always) return false;
      if (PROTECT.has(uiHeader)) return false; // never touch
      const i = uIdx[uiHeader];
      if (i == null) return false;
      const cur = String(vals[i] || '').trim();
      if (always || OVERWRITE || cur === '') { vals[i] = value; return true; }
      return false;
    };

    // Always keep standardized CSU ID
    let changed = false;
    changed |= setIf('CSU ID', csuFmt, true);

    // Fill mapped fields (with normalized values where applicable)
    changed |= setIf('First Name',       first);
    changed |= setIf('Last Name',        last);
    changed |= setIf('Instrument',       instr);
    changed |= setIf('Applied Teacher',  teach);
    changed |= setIf('Accompanist Name', piano);
    changed |= setIf('Recital Length',   length);
    changed |= setIf('Reception',        recept);
    changed |= setIf('Phone Number',     phone);
    changed |= setIf('Email Address',    email);
    changed |= setIf('Recital Type',     rtype);
    changed |= setIf('Recital Date',     reqDt);
    
    // If you want to automatically copy any NEW mapping added to COPY_MAP:
    COPY_MAP.forEach(([srcH, dstH]) => {
      const v = getFromResp(r, srcH);
      changed |= setIf(dstH, v);
    });

    if (changed) {
      ui.getRange(row, 1, 1, uLc).setValues([vals]);
      updated++;
    }
  });

  // Keep RecitalUI sorted by Priority (col A)
  if (created > 0) {
    ui.sort(1); // sorts entire sheet by column A ascending
  }

  SpreadsheetApp.getActive().toast(
    `SyncRecitalUI: added ${created} new row(s), updated ${updated} row(s).`
  );
}
