/* Interactive Excel (.xlsx / .xls) & CSV Importer for BTC Youth Fund Tracker.
   Supports both COYAP Cash Accountability monthly workbooks (including right-side
   Monthly Allowance tables) and standard tabular Excel/CSV spreadsheets. */
(function () {
  const YN = 'Youth Night Offering', JF = 'Joint Fellowship Offering', BB = 'Beginning Balance';
  const AL = 'Allocation from Church Funds', MA = 'Monthly Allowance';
  const MONTH_PREFIXES = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

  const pad2 = n => String(n).padStart(2, '0');
  const numVal = v => {
    if (typeof v === 'number') return isFinite(v) ? v : NaN;
    if (typeof v === 'string') {
      const s = v.replace(/[₱$,\s]/g, '').trim();
      if (!s) return NaN;
      const n = Number(s);
      return isFinite(n) ? n : NaN;
    }
    return NaN;
  };
  const strVal = v => (v == null ? '' : String(v).trim());
  const sheetMonthNum = name => {
    const low = String(name || '').trim().toLowerCase();
    for (let i = 0; i < 12; i++) {
      if (low.startsWith(MONTH_PREFIXES[i])) return i + 1;
    }
    return null;
  };
  const lastSundayDate = (y, mo) => {
    const d = new Date(y, mo, 0);
    while (d.getDay() !== 0) d.setDate(d.getDate() - 1);
    return `${y}-${pad2(mo)}-${pad2(d.getDate())}`;
  };
  const lastFridayDate = (y, mo) => {
    const d = new Date(y, mo, 0);
    while (d.getDay() !== 5) d.setDate(d.getDate() - 1);
    return `${y}-${pad2(mo)}-${pad2(d.getDate())}`;
  };

  function guessCategory(desc) {
    const s = String(desc || '').toLowerCase();
    if (/pansit|snack|food|lunch|dinner|breakfast|softdrink|drink|rice|cake|meal|merienda/.test(s)) return 'Food';
    if (/fabric|linen|material|paper|print|banner|tarpaulin/.test(s)) return 'Materials';
    if (/transport|fare|gas|fuel|van|jeep/.test(s)) return 'Transport';
    if (/venue|rent|hall/.test(s)) return 'Venue';
    if (/honorarium|pastor|speaker|appreciation|gift|prize|price/.test(s)) return 'Honorarium';
    if (/suppl/.test(s)) return 'Supplies';
    return 'Other';
  }

  function parseCellDate(v, fallbackYear, fallbackMonth) {
    if (v instanceof Date && !isNaN(v)) {
      const y = fallbackYear || v.getFullYear();
      const mo = fallbackMonth || (v.getMonth() + 1);
      const day = v.getDate();
      const maxDay = new Date(y, mo, 0).getDate();
      return `${y}-${pad2(mo)}-${pad2(Math.min(day, maxDay))}`;
    }
    if (typeof v === 'number' && v > 20000 && v < 80000) {
      // Excel serial date
      const d = new Date(Math.round((v - 25569) * 86400 * 1000));
      if (!isNaN(d)) {
        const y = fallbackYear || d.getUTCFullYear();
        const mo = fallbackMonth || (d.getUTCMonth() + 1);
        const day = d.getUTCDate();
        const maxDay = new Date(y, mo, 0).getDate();
        return `${y}-${pad2(mo)}-${pad2(Math.min(day, maxDay))}`;
      }
    }
    const s = strVal(v);
    if (!s) return null;
    // YYYY-MM-DD
    let m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (m) return `${fallbackYear || m[1]}-${pad2(+m[2])}-${pad2(+m[3])}`;
    // M/D/YYYY or M-D-YYYY
    m = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})/);
    if (m) {
      const yr = fallbackYear || (+m[3] < 100 ? 2000 + +m[3] : +m[3]);
      return `${yr}-${pad2(+m[1])}-${pad2(+m[2])}`;
    }
    // "Oct. 3" / "Oct 10" / "November 6"
    m = s.match(/([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:[,\s]+(\d{4}))?/);
    if (m) {
      const moIdx = sheetMonthNum(m[1]);
      if (moIdx) {
        const yr = fallbackYear || (m[3] ? +m[3] : new Date().getFullYear());
        const maxDay = new Date(yr, moIdx, 0).getDate();
        return `${yr}-${pad2(moIdx)}-${pad2(Math.min(+m[2], maxDay))}`;
      }
    }
    return null;
  }

  /* Parse one COYAP Cash Accountability sheet (2D array of cells) */
  function parseCoyapSheet(rows, sheetName, targetYear) {
    const entries = [];
    let mo = sheetMonthNum(sheetName);
    // If sheet name isn't a month (e.g. Sheet1), try finding month from "As of ..." or date cells
    if (!mo) {
      for (let r = 0; r < Math.min(rows.length, 25); r++) {
        for (let c = 0; c < Math.min((rows[r] || []).length, 10); c++) {
          const v = rows[r][c];
          const s = strVal(v);
          const mTxt = s.match(/as of\s+([A-Za-z]+)/i);
          if (mTxt && sheetMonthNum(mTxt[1])) { mo = sheetMonthNum(mTxt[1]); break; }
          if (c <= 3 && r >= 10 && r <= 24) {
            const dStr = parseCellDate(v, targetYear, null);
            if (dStr) { mo = +dStr.slice(5, 7); break; }
          }
        }
        if (mo) break;
      }
    }
    if (!mo) return entries;

    const ym = `${targetYear}-${pad2(mo)}`;
    const rowTextLeft = r => (rows[r] || []).slice(0, 12).map(strVal).join(' ').toLowerCase();

    // Locate key row indices on left sheet
    let begIdx = -1, addIdx = -1, ynIdx = -1, jfIdx = -1, allocIdx = -1, totCashIdx = -1, expIdx = -1, netIdx = -1;
    for (let r = 0; r < rows.length; r++) {
      const txt = rowTextLeft(r);
      if (begIdx < 0 && txt.includes('beginning balance')) begIdx = r;
      if (addIdx < 0 && txt.includes('add:') && txt.includes('offering')) addIdx = r;
      if (addIdx >= 0 && r >= addIdx && ynIdx < 0 && txt.includes('youth night offering')) ynIdx = r;
      if (jfIdx < 0 && txt.includes('joint fellowship')) jfIdx = r;
      if (allocIdx < 0 && txt.includes('allocation from church')) allocIdx = r;
      if (totCashIdx < 0 && txt.includes('total cash on hand')) totCashIdx = r;
      if (expIdx < 0 && txt.includes('less:') && txt.includes('expense')) expIdx = r;
      if (expIdx >= 0 && r > expIdx && netIdx < 0 && (txt.includes('net cash on hand') || /^\s*total\b/.test(strVal((rows[r] || [])[0]).toLowerCase()))) netIdx = r;
    }

    // Helper: get numbers in a column range [cMin, cMax]
    const numsInRow = (r, cMin, cMax) => {
      const out = [];
      const row = rows[r] || [];
      for (let c = cMin; c <= Math.min(cMax, row.length - 1); c++) {
        const n = numVal(row[c]);
        if (!isNaN(n) && n > 0) out.push({ col: c, val: n });
      }
      return out;
    };

    // 1. Beginning Balance
    let begBal = 0;
    if (begIdx >= 0) {
      const endBeg = addIdx > begIdx ? addIdx : begIdx + 6;
      for (let r = begIdx; r < endBeg; r++) {
        const nums = numsInRow(r, 2, 11);
        if (nums.length) begBal = nums[nums.length - 1].val;
      }
      if (begBal > 0) {
        entries.push({ type: 'income', fund: 'general', date: `${ym}-01`, amount: begBal, category: BB, notes: 'Cash on Hand - Beginning Balance', sheet: sheetName });
      }
    }

    // 2. Youth Night Offerings
    let totalYN = 0;
    if (ynIdx >= 0) {
      const endYn = jfIdx > ynIdx ? jfIdx : Math.min(rows.length, ynIdx + 10);
      for (let r = ynIdx + 1; r < endYn; r++) {
        const row = rows[r] || [];
        let foundDate = null;
        for (let c = 0; c <= Math.min(4, row.length - 1); c++) {
          const d = parseCellDate(row[c], targetYear, mo);
          if (d) { foundDate = d; break; }
        }
        if (!foundDate) continue;
        const nums = numsInRow(r, 2, 8);
        if (nums.length) {
          const amt = nums[0].val;
          totalYN += amt;
          entries.push({ type: 'income', fund: 'general', date: foundDate, amount: amt, category: YN, notes: YN, sheet: sheetName });
        }
      }
    }

    // Total Cash on Hand (reference value to distinguish subtotals from real entries)
    let totalCashCell = 0;
    if (totCashIdx >= 0) {
      const nums = numsInRow(totCashIdx, 2, 11);
      if (nums.length) totalCashCell = nums[nums.length - 1].val;
    }

    // 3. Joint Fellowship Offering
    let totalJF = 0;
    if (jfIdx >= 0) {
      const subNums = numsInRow(jfIdx, 2, 7);
      const totNums = numsInRow(jfIdx, 8, 11);
      if (subNums.length) {
        totalJF = subNums[subNums.length - 1].val;
      } else if (totNums.length) {
        const cand = totNums[0].val;
        // Avoid treating the Additions subtotal mirror (= totalYN) as Joint Fellowship
        const isSubtotalMirror = totalYN > 0 && Math.abs(cand - totalYN) < 0.01 &&
          (totalCashCell === 0 || Math.abs(begBal + totalYN - totalCashCell) < 1.5);
        if (!isSubtotalMirror) totalJF = cand;
      }
      if (totalJF > 0) {
        entries.push({ type: 'income', fund: 'general', date: lastSundayDate(targetYear, mo), amount: totalJF, category: JF, notes: JF, sheet: sheetName });
      }
    }

    // 4. Allocation from Church Funds (left side only if it's an actual allocation, not the Offerings subtotal)
    if (allocIdx >= 0) {
      const subNums = numsInRow(allocIdx, 2, 7);
      const totNums = numsInRow(allocIdx, 8, 11);
      let leftAlloc = 0;
      if (subNums.length) {
        leftAlloc = subNums[subNums.length - 1].val;
      } else if (totNums.length) {
        const cand = totNums[0].val;
        const sumOfferings = totalYN + totalJF;
        const isSubtotalMirror = (sumOfferings > 0 && (Math.abs(cand - sumOfferings) < 0.01 || Math.abs(cand - totalYN) < 0.01)) ||
          (totalCashCell > 0 && Math.abs(begBal + sumOfferings - totalCashCell) < 1.5);
        if (!isSubtotalMirror) leftAlloc = cand;
      }
      if (leftAlloc > 0) {
        entries.push({ type: 'income', fund: 'allowance', date: `${ym}-01`, amount: leftAlloc, category: AL, notes: 'Allocation from Church Funds', sheet: sheetName });
      }
    }

    // 5. Less: Expenses (left side)
    if (expIdx >= 0) {
      const endExp = netIdx > expIdx ? netIdx : Math.min(rows.length, expIdx + 10);
      for (let r = expIdx; r < endExp; r++) {
        const row = rows[r] || [];
        const txt0 = strVal(row[0]).toLowerCase();
        if (r > expIdx && (txt0.includes('net cash') || txt0 === 'total' || txt0.includes('prepared by'))) break;
        const nums = numsInRow(r, 2, 11);
        if (!nums.length) continue;
        const amt = nums[nums.length - 1].val;
        // Gather description text & optional date from cols 1..8
        let expDate = lastFridayDate(targetYear, mo);
        const parts = [];
        for (let c = 1; c <= Math.min(8, row.length - 1); c++) {
          const cell = row[c];
          if (cell == null) continue;
          const d = parseCellDate(cell, targetYear, mo);
          if (d) { expDate = d; continue; }
          if (typeof cell === 'string') {
            const s = cell.trim();
            if (s && isNaN(numVal(s)) && !/^less:\s*expenses?$/i.test(s)) parts.push(s);
          }
        }
        const desc = parts.join(' - ') || 'Expense';
        entries.push({ type: 'expense', fund: 'general', date: expDate, amount: amt, category: guessCategory(desc), notes: desc, sheet: sheetName });
      }
    }

    // 6. Right-side Monthly Allowance block (columns >= 12)
    let allowStart = -1, allowCol = -1;
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r] || [];
      for (let c = 12; c < row.length; c++) {
        const s = strVal(row[c]).toLowerCase();
        if (s.includes('allowance')) { allowStart = r; allowCol = c; break; }
      }
      if (allowStart >= 0) break;
    }
    if (allowStart >= 0) {
      let inAllowExp = false;
      let allowAllocAdded = false;
      const maxR = Math.min(rows.length, allowStart + 18);
      for (let r = allowStart; r < maxR; r++) {
        const row = rows[r] || [];
        const rightText = row.slice(Math.max(12, allowCol - 1)).map(strVal).join(' ').toLowerCase();
        if (r > allowStart && (rightText.includes('net cash') || /\bnet\s*:/.test(rightText) || /^\s*total\s*$/.test(strVal(row[allowCol]).toLowerCase()))) break;
        if (rightText.includes('less') && (rightText.includes('expense') || rightText.includes('expence'))) {
          inAllowExp = true;
        }
        const nums = numsInRow(r, allowCol, row.length - 1);
        if (!nums.length) continue;
        const amt = nums[nums.length - 1].val;

        if (!inAllowExp) {
          if (!allowAllocAdded && (rightText.includes('monthly') || rightText.includes('total') || rightText.includes('cash on hand') || r === allowStart)) {
            allowAllocAdded = true;
            entries.push({ type: 'income', fund: 'allowance', date: `${ym}-01`, amount: amt, category: AL, notes: 'Monthly Allowance from Church', sheet: sheetName });
          }
        } else {
          // Expense from allowance
          const descParts = [];
          for (let c = allowCol; c < row.length; c++) {
            const cell = row[c];
            if (typeof cell === 'string') {
              const s = cell.trim();
              if (s && isNaN(numVal(s)) && !/less.*expen/i.test(s)) descParts.push(s);
            }
          }
          const desc = descParts.join(' - ') || 'Monthly Allowance Expense';
          entries.push({ type: 'expense', fund: 'allowance', date: lastFridayDate(targetYear, mo), amount: amt, category: MA, notes: desc, sheet: sheetName });
        }
      }
    }

    return entries;
  }

  /* Parse a generic tabular sheet (headers like Date, Type, Category, Amount, Notes, Fund) */
  function parseGenericSheet(rows, sheetName, targetYear) {
    const entries = [];
    let hdrIdx = -1, colDate = -1, colAmt = -1, colType = -1, colCat = -1, colNotes = -1, colFund = -1;
    for (let r = 0; r < Math.min(rows.length, 15); r++) {
      const row = (rows[r] || []).map(v => strVal(v).toLowerCase());
      const dIdx = row.findIndex(h => h === 'date' || h.includes('date'));
      const aIdx = row.findIndex(h => h === 'amount' || h.includes('amount'));
      if (dIdx >= 0 && aIdx >= 0) {
        hdrIdx = r;
        colDate = dIdx;
        colAmt = aIdx;
        colType = row.findIndex(h => h === 'type');
        colCat = row.findIndex(h => h.includes('cat'));
        colNotes = row.findIndex(h => h.includes('note') || h.includes('desc') || h.includes('particular'));
        colFund = row.findIndex(h => h.includes('fund'));
        break;
      }
    }
    if (hdrIdx < 0) return entries;

    for (let r = hdrIdx + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const dStr = parseCellDate(row[colDate], null, null);
      const rawAmt = numVal(row[colAmt]);
      if (!dStr || isNaN(rawAmt) || rawAmt === 0) continue;
      let type = colType >= 0 ? strVal(row[colType]).toLowerCase() : '';
      if (type !== 'income' && type !== 'expense') type = rawAmt < 0 ? 'expense' : 'income';
      const amount = Math.abs(rawAmt);
      const notes = colNotes >= 0 ? strVal(row[colNotes]) : '';
      const fundStr = colFund >= 0 ? strVal(row[colFund]).toLowerCase() : '';
      const category = (colCat >= 0 && strVal(row[colCat])) || (type === 'income' ? 'Offering' : guessCategory(notes));
      const fund = (fundStr.includes('allow') || category === AL || category === MA) ? 'allowance' : 'general';
      entries.push({ type, fund, date: dStr, amount, category, notes, sheet: sheetName });
    }
    return entries;
  }

  function parseCsvText(text) {
    const rows = [];
    let row = [], cur = '', inQ = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (inQ) {
        if (ch === '"' && text[i + 1] === '"') { cur += '"'; i++; }
        else if (ch === '"') inQ = false;
        else cur += ch;
      } else if (ch === '"') {
        inQ = true;
      } else if (ch === ',' || ch === '\t') {
        row.push(cur); cur = '';
      } else if (ch === '\r') {
        continue;
      } else if (ch === '\n') {
        row.push(cur); rows.push(row); row = []; cur = '';
      } else {
        cur += ch;
      }
    }
    if (cur.length || row.length) { row.push(cur); rows.push(row); }
    return rows;
  }

  /* Parse a SheetJS workbook object into transaction entries */
  function parseWorkbook(wb, targetYear) {
    const all = [];
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
      const flatText = rows.slice(0, 25).flat().map(strVal).join(' ').toLowerCase();
      const isCoyap = flatText.includes('cash accountability') || flatText.includes('youth night offering') || flatText.includes('beginning balance');
      const parsed = isCoyap ? parseCoyapSheet(rows, name, targetYear) : parseGenericSheet(rows, name, targetYear);
      all.push(...parsed);
    }
    return all;
  }
  window._parseExcelWorkbook = parseWorkbook;
  window._parseCoyapSheet = parseCoyapSheet;

  /* Built-in 2026 sample entries (preserved for quick one-click demo/import) */
  const E = [];
  const inc = (date, amount, category, notes = '', fund = 'general') => amount > 0 && E.push({ type: 'income', fund, date, amount, category, notes });
  const exp = (date, amount, notes, fund = 'general', category = 'Other') => amount > 0 && E.push({ type: 'expense', fund, date, amount, category: fund === 'allowance' ? MA : category, notes });
  const bb = (d, a) => inc(d, a, BB, 'Cash on Hand - Beginning Balance');
  const allow = (d, a) => inc(d, a, AL, 'Monthly allowance from church', 'allowance');
  bb('2026-01-01', 3583); inc('2026-01-23', 290, YN, YN); inc('2026-01-30', 165, YN, YN); inc('2026-01-25', 176, JF, JF); exp('2026-01-30', 500, 'Pansit', 'general', 'Food');
  bb('2026-02-01', 3715); inc('2026-02-06', 100, YN, YN); inc('2026-02-13', 104, YN, YN); inc('2026-02-20', 90, YN, YN); inc('2026-02-27', 75, YN, YN); inc('2026-02-22', 170, JF, JF); allow('2026-02-01', 1000); exp('2026-02-27', 500, 'Less expenses', 'allowance');
  bb('2026-03-01', 3754); inc('2026-03-06', 270, YN, YN); inc('2026-03-20', 170, YN, YN); inc('2026-03-29', 414, JF, JF); allow('2026-03-01', 1000); exp('2026-03-27', 480, 'Less expenses', 'allowance'); exp('2026-03-27', 520, 'Less expenses', 'allowance');
  bb('2026-04-01', 4608); inc('2026-04-10', 370, YN, YN); inc('2026-04-17', 275, YN, YN); inc('2026-04-24', 208, YN, YN); inc('2026-04-26', 220, JF, JF); exp('2026-04-26', 1568, 'Church Baptism - pansit/softdrinks', 'general', 'Food'); exp('2026-04-26', 2000, 'Church project', 'general', 'Other'); allow('2026-04-01', 1000); exp('2026-04-24', 820, 'Pansit', 'allowance'); exp('2026-04-24', 138, 'White linen', 'allowance');
  bb('2026-05-01', 2113);
  bb('2026-06-01', 2113); inc('2026-06-19', 230, YN, YN); inc('2026-06-26', 215, YN, YN); exp('2026-06-26', 200, 'White fabric', 'general', 'Materials'); exp('2026-06-26', 400, 'Snacks', 'general', 'Food'); allow('2026-06-01', 1000); exp('2026-06-26', 350, 'Less expenses', 'allowance'); exp('2026-06-26', 239, 'Less expenses', 'allowance');
  bb('2026-07-01', 1953); inc('2026-07-03', 85, YN, YN); inc('2026-07-10', 103, YN, YN); inc('2026-07-17', 81, YN, YN); inc('2026-07-24', 127, YN, YN); inc('2026-07-26', 69, JF, JF); exp('2026-07-31', 84, 'Expense', 'general', 'Other'); allow('2026-07-01', 1000); allow('2026-07-01', 84); exp('2026-07-31', 707, 'Pansit', 'allowance'); exp('2026-07-31', 150, 'Rice', 'allowance'); exp('2026-07-31', 227, 'Prizes', 'allowance');
  bb('2026-08-01', 2334); inc('2026-08-07', 110, YN, YN); inc('2026-08-14', 120, YN, YN); inc('2026-08-21', 112, YN, YN); inc('2026-08-28', 293, YN, YN); inc('2026-08-30', 190, JF, JF); exp('2026-08-28', 972.1, 'Snacks', 'general', 'Food'); allow('2026-08-01', 1000); exp('2026-08-28', 556, 'Pansit', 'allowance'); exp('2026-08-28', 95, 'Rice', 'allowance'); exp('2026-08-28', 95.25, 'Price', 'allowance');
  bb('2026-09-01', 2440); inc('2026-09-04', 62, YN, YN); inc('2026-09-11', 92, YN, YN); allow('2026-09-01', 1000); exp('2026-09-25', 250, 'Less expenses', 'allowance');
  allow('2026-10-01', 1000); exp('2026-10-02', 600, 'Less expenses', 'allowance'); exp('2026-10-09', 304, 'Less expenses', 'allowance');
  allow('2026-11-01', 1000); exp('2026-11-27', 831, 'Joint Lunch', 'allowance');
  window.EXCEL_IMPORT_ENTRIES = E;

  /* ---------- Interactive Import Excel Modal ---------- */
  window.importExcelReport = function () {
    let loadedFile = null;
    let loadedWb = null;
    let loadedCsvRows = null;
    let parsedEntries = [];
    let targetYear = new Date().getFullYear();

    const dupKey = e => [e.type, e.date, Number(e.amount).toFixed(2), e.category].join('|');
    const existingKeys = () => new Set(S.transactions.filter(t => t.status !== 'void').map(dupKey));

    function reparse() {
      if (loadedWb) {
        parsedEntries = parseWorkbook(loadedWb, targetYear);
      } else if (loadedCsvRows) {
        const flatText = loadedCsvRows.slice(0, 25).flat().map(strVal).join(' ').toLowerCase();
        const isCoyap = flatText.includes('cash accountability') || flatText.includes('youth night offering') || flatText.includes('beginning balance');
        parsedEntries = isCoyap
          ? parseCoyapSheet(loadedCsvRows, loadedFile?.name || 'Sheet', targetYear)
          : parseGenericSheet(loadedCsvRows, loadedFile?.name || 'CSV', targetYear);
      }
      renderPreview();
    }

    async function handleFile(file) {
      if (!file) return;
      loadedFile = file;
      loadedWb = null;
      loadedCsvRows = null;
      const ymMatch = file.name.match(/\b(20\d{2})\b/);
      if (ymMatch) {
        targetYear = +ymMatch[1];
        const yrInp = document.getElementById('imp-year');
        if (yrInp) yrInp.value = targetYear;
      }

      const statusEl = document.getElementById('imp-status');
      if (statusEl) statusEl.innerHTML = `<span class="mute">Reading <b>${esc(file.name)}</b>…</span>`;

      try {
        const lower = file.name.toLowerCase();
        if (lower.endsWith('.csv') || lower.endsWith('.tsv') || lower.endsWith('.txt')) {
          loadedCsvRows = parseCsvText(await file.text());
        } else {
          if (typeof XLSX === 'undefined') {
            throw new Error('Excel library (SheetJS) did not load. Please check your internet connection or save the file as .CSV and import that.');
          }
          const buf = await file.arrayBuffer();
          loadedWb = XLSX.read(buf, { type: 'array', cellDates: true });
        }
        reparse();
      } catch (err) {
        if (statusEl) statusEl.innerHTML = `<div class="alert" style="margin:0">⚠️ Could not read file: ${esc(err.message || err)}</div>`;
      }
    }

    function renderPreview() {
      const box = document.getElementById('imp-preview');
      const statusEl = document.getElementById('imp-status');
      const submitBtn = document.getElementById('imp-confirm-btn');
      if (!box) return;

      const have = existingKeys();
      const annotated = parsedEntries.map((e, i) => ({ ...e, idx: i, isDup: have.has(dupKey(e)) }));
      const freshCount = annotated.filter(x => !x.isDup).length;
      const dupCount = annotated.length - freshCount;

      if (!annotated.length) {
        if (statusEl) statusEl.innerHTML = loadedFile ? `<div class="alert warn" style="margin:0">No transactions were found in <b>${esc(loadedFile.name)}</b>. Try checking the Target Year or use a COYAP report / Date+Amount table.</div>` : '';
        box.innerHTML = '';
        if (submitBtn) { submitBtn.disabled = true; submitBtn.textContent = '📥 Import Selected'; }
        return;
      }

      if (statusEl) {
        statusEl.innerHTML = `<div class="ms-banner ${freshCount ? 'ok' : 'warn'}" style="margin:0">
          <span>📄 <b>${esc(loadedFile ? loadedFile.name : 'Built-in 2026 Report')}</b> — found <b>${annotated.length}</b> entries (<b>${freshCount} new</b>${dupCount ? `, ${dupCount} already saved` : ''}).</span>
        </div>`;
      }

      box.innerHTML = `
        <div style="display:flex;justify-content:space-between;align-items:center;margin:10px 0 6px;flex-wrap:wrap;gap:8px">
          <div style="display:flex;gap:8px">
            <button type="button" class="btn sm ghost" id="imp-sel-new">Select new (${freshCount})</button>
            <button type="button" class="btn sm ghost" id="imp-sel-all">Select all (${annotated.length})</button>
            <button type="button" class="btn sm ghost" id="imp-sel-none">Clear</button>
          </div>
          <span class="mute" style="font-size:12px">Uncheck any row you don't want to import</span>
        </div>
        <div style="max-height:320px;overflow:auto;border:1px solid var(--line);border-radius:12px">
          <table>
            <tr><th style="width:32px"></th><th>Date</th><th>Type</th><th>Category</th><th>Fund</th><th>Notes</th><th style="text-align:right">Amount</th><th>Status</th></tr>
            ${annotated.map(e => `<tr style="${e.isDup ? 'opacity:.55' : ''}">
              <td><input type="checkbox" class="imp-chk" data-idx="${e.idx}" ${e.isDup ? '' : 'checked'} style="width:auto"></td>
              <td style="white-space:nowrap">${e.date}</td>
              <td>${e.type}</td>
              <td><span class="tag">${esc(e.category)}</span></td>
              <td><span class="fund-badge-pill ${e.fund === 'allowance' ? 'fund-badge-allowance' : 'fund-badge-general'}">${e.fund === 'allowance' ? '💳 Allowance' : '🏢 General'}</span></td>
              <td class="mute">${esc(e.notes || '—')}</td>
              <td class="${e.type === 'income' ? 'pos' : 'neg'}" style="text-align:right;white-space:nowrap">${e.type === 'income' ? '+' : '−'}${money(e.amount)}</td>
              <td>${e.isDup ? '<span class="tag">Duplicate</span>' : '<span class="tag ok">New</span>'}</td>
            </tr>`).join('')}
          </table>
        </div>`;

      const updateCount = () => {
        const n = box.querySelectorAll('.imp-chk:checked').length;
        if (submitBtn) {
          submitBtn.disabled = n === 0;
          submitBtn.textContent = n ? `📥 Import ${n} Selected` : '📥 Select entries to import';
        }
      };
      box.querySelectorAll('.imp-chk').forEach(c => c.onchange = updateCount);
      document.getElementById('imp-sel-new').onclick = () => {
        box.querySelectorAll('.imp-chk').forEach(c => { c.checked = !annotated[+c.dataset.idx].isDup; });
        updateCount();
      };
      document.getElementById('imp-sel-all').onclick = () => {
        box.querySelectorAll('.imp-chk').forEach(c => { c.checked = true; });
        updateCount();
      };
      document.getElementById('imp-sel-none').onclick = () => {
        box.querySelectorAll('.imp-chk').forEach(c => { c.checked = false; });
        updateCount();
      };
      updateCount();
    }

    openModal(`
      <h2>📥 Import Data from Excel / CSV</h2>
      <p class="mute" style="margin-bottom:12px">Upload any <b>.xlsx</b>, <b>.xls</b>, or <b>.csv</b> file — supports your monthly COYAP Cash Accountability workbooks (with all monthly tabs &amp; Allowance tables) as well as standard Date/Amount spreadsheets.</p>

      <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:12px">
        <label style="display:flex;align-items:center;gap:8px;font-weight:600;color:var(--text)">Target Year:
          <input type="number" id="imp-year" value="${targetYear}" min="2020" max="2035" style="width:96px;padding:6px 10px">
        </label>
        <span class="mute" style="font-size:12px">Fixes sheets that still say "As of 2024/2025" inside a new year's workbook</span>
      </div>

      <label id="imp-dropzone" class="imp-dropzone">
        <input type="file" id="imp-file-input" accept=".xlsx,.xls,.csv,.tsv" style="display:none">
        <div style="font-size:28px;margin-bottom:4px">📂</div>
        <div style="font-weight:700;font-size:15px;color:var(--text)">Click to choose an Excel or CSV file</div>
        <div class="mute" style="font-size:12px;margin-top:2px">or drag and drop your .xlsx / .xls / .csv file here</div>
      </label>

      <div id="imp-status" style="margin-top:12px"></div>
      <div id="imp-preview"></div>

      <div class="row" style="margin-top:16px;flex-wrap:wrap;gap:10px">
        <button type="button" class="btn sm ghost" id="imp-builtin-btn" title="Load built-in 2026 financial report data">Load built-in 2026 data</button>
        <div style="display:flex;gap:10px">
          <button type="button" class="btn ghost" onclick="closeModal()">Cancel</button>
          <button type="button" class="btn" id="imp-confirm-btn" disabled style="background:linear-gradient(135deg,#06b6d4,#3b82f6)">📥 Import Selected</button>
        </div>
      </div>
    `, 'landscape');

    const dz = document.getElementById('imp-dropzone');
    const fi = document.getElementById('imp-file-input');
    const yr = document.getElementById('imp-year');

    fi.onchange = () => handleFile(fi.files[0]);
    dz.ondragover = e => { e.preventDefault(); dz.classList.add('dragover'); };
    dz.ondragleave = () => dz.classList.remove('dragover');
    dz.ondrop = e => {
      e.preventDefault();
      dz.classList.remove('dragover');
      if (e.dataTransfer?.files?.[0]) handleFile(e.dataTransfer.files[0]);
    };
    yr.onchange = () => {
      const v = parseInt(yr.value, 10);
      if (v >= 2000 && v <= 2100) {
        targetYear = v;
        if (loadedWb || loadedCsvRows) reparse();
      }
    };
    document.getElementById('imp-builtin-btn').onclick = () => {
      loadedFile = null;
      loadedWb = null;
      loadedCsvRows = null;
      parsedEntries = E.map(x => ({ ...x }));
      renderPreview();
    };
    document.getElementById('imp-confirm-btn').onclick = async () => {
      const box = document.getElementById('imp-preview');
      const selected = [...box.querySelectorAll('.imp-chk:checked')].map(c => parsedEntries[+c.dataset.idx]).filter(Boolean);
      if (!selected.length) return;
      const btn = document.getElementById('imp-confirm-btn');
      btn.disabled = true;
      btn.textContent = 'Importing…';
      await window.saveBatchOfferings(selected);
      const dates = selected.map(x => x.date).filter(Boolean).sort();
      if (dates.length && typeof setActiveMonth === 'function') {
        setActiveMonth(dates[dates.length - 1].slice(0, 7));
      }
      if (typeof showToast === 'function') showToast(`✔ Imported ${selected.length} entries from Excel`);
    };
  };
})();
