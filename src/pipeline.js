import * as XLSX from 'xlsx';

/**
 * MinnARK Revenue Split Pipeline - JavaScript Port
 * Mirrors pipeline.py logic exactly.
 */

function parseDate(val) {
  if (val == null) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'number') {
    // Excel serial date
    const d = XLSX.SSF.parse_date_code(val);
    if (d) return new Date(d.y, d.m - 1, d.d);
  }
  if (typeof val === 'string') {
    const s = val.trim();
    // Try common formats
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function findCol(headers, ...names) {
  const lowerNames = names.map(n => n.toLowerCase());
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] == null) continue;
    const h = String(headers[i]).trim().toLowerCase();
    if (lowerNames.includes(h)) return i;
  }
  return null;
}

function getMonth(date) {
  return date.getFullYear() * 100 + (date.getMonth() + 1);
}

function getMonthName(ym) {
  const months = ['', 'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'];
  return `${months[ym % 100]} ${Math.floor(ym / 100)}`;
}

/**
 * Detect file type from sheet names.
 * Returns 'jpm' | 'domestic' | null
 */
export function detectFileType(workbook) {
  const names = workbook.SheetNames.map(n => n.trim());
  if (names.some(n => n === 'TOTAL DI' || n === 'SUMMARY')) return 'jpm';
  if (names.some(n => n.includes('Payment Details'))) return 'domestic';
  return null;
}

/**
 * Read TOTAL DI tab from a JPM workbook.
 * Returns { rows: [{team, program, amount}], month: string }
 */
export function readTotalDI(workbook, filename) {
  const ws = workbook.Sheets['TOTAL DI'];
  if (!ws) throw new Error(`No "TOTAL DI" sheet found in ${filename}`);

  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (data.length < 2) throw new Error('TOTAL DI sheet is empty');

  const headers = data[0];
  const teamCol = findCol(headers, 'Team');
  const progCol = findCol(headers, 'Program');
  const paidCol = findCol(headers, 'Paid Per Item', ' Paid Per Item ');
  const netCol = findCol(headers, 'Payment Amount', 'Discount Net Amount');

  if (teamCol == null) throw new Error(`No 'Team' column in TOTAL DI headers: ${headers.join(', ')}`);

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const rv = data[i];
    if (!rv || rv.slice(0, 5).every(v => v == null)) continue;

    const team = rv[teamCol] ? String(rv[teamCol]).trim() : null;
    const program = progCol != null && rv[progCol] ? String(rv[progCol]).trim() : null;

    let amount = null;
    if (paidCol != null) amount = rv[paidCol];
    if (amount == null && netCol != null) amount = rv[netCol];
    if (amount == null) continue;

    amount = Number(amount);
    if (isNaN(amount)) continue;

    if (!team) {
      // Skip total/subtotal rows
      const hasLabel = rv.slice(0, 8).some(v =>
        v != null && ['TOTAL', 'SUBTOTAL', 'BLACKFIN', 'MIZAR'].includes(String(v).trim().toUpperCase())
      );
      if (hasLabel) continue;
      continue; // skip no-team rows
    }

    rows.push({ team, program: program || 'UNKNOWN', amount });
  }

  // Detect month from SUMMARY or filename
  let month = filename;
  return { rows, filename };
}

/**
 * Read Domestic Payments workbook.
 * Returns [{date, team, program, amount}]
 */
export function readDomesticPayments(workbook) {
  // Find the payment details sheet
  // Prefer 2025 sheet over 2024
  const sheetName = workbook.SheetNames.find(n => n.includes('2025_Payment Details'))
    || workbook.SheetNames.find(n => n.includes('Payment Details'));
  if (!sheetName) throw new Error('No Payment Details sheet found');

  const ws = workbook.Sheets[sheetName];
  // Limit read to columns A-AH (0-33) to avoid processing 16K+ empty columns
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  range.e.c = Math.min(range.e.c, 33); // cap at column AH
  ws['!ref'] = XLSX.utils.encode_range(range);
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  if (data.length < 2) throw new Error('Domestic Payments sheet is empty');

  console.log('[DO] Sheet:', sheetName, '| Rows:', data.length, '| Row1 cols:', data[1]?.length);
  console.log('[DO] Row1 sample — col2:', data[1]?.[2], '| col27:', data[1]?.[27], '| col32:', data[1]?.[32]);

  const rows = [];
  let lastDate = null;
  let debugNoDate = 0, debugNoTeam = 0, debugNoPaid = 0;

  for (let i = 1; i < data.length; i++) {
    const rv = data[i];
    if (!rv) continue;

    // Payment Date = col C (index 2)
    const pd = parseDate(rv[2]);
    if (pd) lastDate = pd;

    // Team = col AB (index 27)
    const teamRaw = rv[27];
    const team = teamRaw ? String(teamRaw).trim() : null;

    // Program = col AH (index 33)
    const progRaw = rv[33];
    const program = progRaw ? String(progRaw).trim() : null;

    // Paid Per Item = col AG (index 32)
    const paid = rv[32];

    if (!lastDate) { debugNoDate++; continue; }
    if (!team) { debugNoTeam++; continue; }
    if (paid == null) { debugNoPaid++; continue; }

    const amount = Number(paid);
    if (!isNaN(amount)) {
      rows.push({
        date: lastDate,
        team,
        program: program || 'UNKNOWN',
        amount,
      });
    }
  }

  console.log('[DO] Result:', rows.length, 'valid rows | noDate:', debugNoDate, '| noTeam:', debugNoTeam, '| noPaid:', debugNoPaid);
  if (rows.length > 0) {
    const byMonth = {};
    for (const r of rows) {
      const ym = r.date.getFullYear() * 100 + (r.date.getMonth() + 1);
      byMonth[ym] = (byMonth[ym] || 0) + 1;
    }
    console.log('[DO] By month:', byMonth);
  }

  return rows;
}

/**
 * Aggregate rows by team and program.
 */
function aggregate(rows) {
  const result = {};
  for (const r of rows) {
    if (!result[r.team]) result[r.team] = {};
    result[r.team][r.program] = (result[r.team][r.program] || 0) + r.amount;
  }
  return result;
}

/**
 * Process all uploaded files.
 * @param {Array} jpmFiles - [{workbook, filename}]
 * @param {Object|null} domesticFile - {workbook, filename} or null
 * @returns {Array} monthly results
 */
export function processFiles(jpmFiles, domesticFile) {
  // Load all DO data
  let allDoRows = [];
  if (domesticFile) {
    allDoRows = readDomesticPayments(domesticFile.workbook);
  }

  // Group DO rows by month
  const doByMonth = {};
  for (const r of allDoRows) {
    const ym = getMonth(r.date);
    if (!doByMonth[ym]) doByMonth[ym] = [];
    doByMonth[ym].push(r);
  }

  const results = [];

  for (const { workbook, filename } of jpmFiles) {
    const { rows: diRows } = readTotalDI(workbook, filename);
    const diAgg = aggregate(diRows);

    // Determine which months of DO data to include
    // We need to figure out the month from the file. Try reading SUMMARY or use DO data months.
    // For each JPM file, find unique months in DO data and create a result per month.

    // Try to detect months from filename or SUMMARY tab
    const monthsInFile = detectMonthsFromJPM(workbook, filename);

    for (const { month, monthKey } of monthsInFile) {
      const doRows = doByMonth[monthKey] || [];
      const doAgg = aggregate(doRows);

      const actual = {
        di_blackfin: sum(diAgg['Blackfin']),
        di_mizar: sum(diAgg['Mizar']),
        do_blackfin: sum(doAgg['Blackfin']),
        do_mizar: sum(doAgg['Mizar']),
      };
      actual.di_total = actual.di_blackfin + actual.di_mizar;
      actual.do_total = actual.do_blackfin + actual.do_mizar;
      actual.grand_total = actual.di_total + actual.do_total;

      results.push({
        month,
        monthKey,
        filename,
        actual,
        di_blackfin_programs: roundObj(diAgg['Blackfin'] || {}),
        di_mizar_programs: roundObj(diAgg['Mizar'] || {}),
        do_blackfin_programs: roundObj(doAgg['Blackfin'] || {}),
        do_mizar_programs: roundObj(doAgg['Mizar'] || {}),
        di_row_count: diRows.length,
        do_row_count: doRows.length,
      });
    }
  }

  // Sort by monthKey
  results.sort((a, b) => a.monthKey - b.monthKey);
  return results;
}

function detectMonthsFromJPM(workbook, filename) {
  // Try to detect month(s) from the TOTAL DI data dates, or fallback to filename parsing
  // Each JPM file represents ONE month of DI data
  // Try reading dates from TOTAL DI
  const ws = workbook.Sheets['TOTAL DI'];
  if (!ws) return [{ month: filename, monthKey: 0 }];

  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  const headers = data[0] || [];
  const dateCol = findCol(headers, 'Payment Date', 'Discount Start Date');

  const months = new Set();
  if (dateCol != null) {
    for (let i = 1; i < data.length; i++) {
      const d = parseDate(data[i]?.[dateCol]);
      if (d) months.add(getMonth(d));
    }
  }

  if (months.size === 0) {
    // Fallback: parse from filename
    const monthNames = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
    const match = filename.match(/(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/i);
    const yearMatch = filename.match(/(\d{4})/);
    if (match && yearMatch) {
      const m = monthNames[match[1].toLowerCase()];
      const y = parseInt(yearMatch[1]);
      // For files like "Oct-Nov", the TOTAL DI tab is for the FIRST month mentioned
      return [{ month: getMonthName(y * 100 + m), monthKey: y * 100 + m }];
    }
    return [{ month: filename, monthKey: 0 }];
  }

  // If multiple months detected, use the most common one (the primary month)
  // Actually, TOTAL DI contains data for one logical month per file
  // Just use the most frequent month
  const counts = {};
  if (dateCol != null) {
    for (let i = 1; i < data.length; i++) {
      const d = parseDate(data[i]?.[dateCol]);
      if (d) {
        const ym = getMonth(d);
        counts[ym] = (counts[ym] || 0) + 1;
      }
    }
  }

  const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  const primaryMonth = parseInt(sorted[0][0]);
  return [{ month: getMonthName(primaryMonth), monthKey: primaryMonth }];
}

function sum(obj) {
  if (!obj) return 0;
  return Object.values(obj).reduce((a, b) => a + b, 0);
}

function roundObj(obj) {
  const r = {};
  for (const [k, v] of Object.entries(obj)) {
    r[k] = Math.round(v * 100) / 100;
  }
  return r;
}

/**
 * Export results to CSV string
 */
export function exportCSV(results) {
  const lines = ['Month,Category,Team,Program,Amount'];
  for (const r of results) {
    for (const [prog, amt] of Object.entries(r.di_blackfin_programs)) {
      lines.push(`${r.month},DI,Blackfin,${prog},${amt.toFixed(2)}`);
    }
    for (const [prog, amt] of Object.entries(r.di_mizar_programs)) {
      lines.push(`${r.month},DI,Mizar,${prog},${amt.toFixed(2)}`);
    }
    for (const [prog, amt] of Object.entries(r.do_blackfin_programs)) {
      lines.push(`${r.month},DO,Blackfin,${prog},${amt.toFixed(2)}`);
    }
    for (const [prog, amt] of Object.entries(r.do_mizar_programs)) {
      lines.push(`${r.month},DO,Mizar,${prog},${amt.toFixed(2)}`);
    }
  }
  return lines.join('\n');
}

/**
 * Export results to Excel workbook (as ArrayBuffer)
 */
export function exportExcel(results) {
  const wb = XLSX.utils.book_new();

  // Summary sheet
  const summaryData = [['Month', '', 'Blackfin DI', 'Mizar DI', 'Blackfin DO', 'Mizar DO', 'Grand Total']];
  for (const r of results) {
    summaryData.push([
      r.month, '',
      r.actual.di_blackfin, r.actual.di_mizar,
      r.actual.do_blackfin, r.actual.do_mizar,
      r.actual.grand_total,
    ]);
  }
  // Totals row
  if (results.length > 1) {
    const totals = ['TOTAL', ''];
    for (const key of ['di_blackfin', 'di_mizar', 'do_blackfin', 'do_mizar', 'grand_total']) {
      totals.push(results.reduce((s, r) => s + r.actual[key], 0));
    }
    summaryData.push(totals);
  }
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryData), 'Summary');

  // Detail sheet
  const detailData = [['Month', 'Category', 'Team', 'Program', 'Amount']];
  for (const r of results) {
    const addRows = (cat, team, progs) => {
      for (const [prog, amt] of Object.entries(progs)) {
        detailData.push([r.month, cat, team, prog, amt]);
      }
    };
    addRows('DI', 'Blackfin', r.di_blackfin_programs);
    addRows('DI', 'Mizar', r.di_mizar_programs);
    addRows('DO', 'Blackfin', r.do_blackfin_programs);
    addRows('DO', 'Mizar', r.do_mizar_programs);
  }
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(detailData), 'Detail');

  return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
}

/**
 * Read an uploaded file as XLSX workbook
 */
export function readFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        resolve(workbook);
      } catch (err) {
        reject(new Error(`Failed to parse ${file.name}: ${err.message}`));
      }
    };
    reader.onerror = () => reject(new Error(`Failed to read ${file.name}`));
    reader.readAsArrayBuffer(file);
  });
}
