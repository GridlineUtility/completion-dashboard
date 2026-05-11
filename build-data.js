// build-data.js — run with: node build-data.js
// Parses all CompletionDetailReport XLS files and writes data.json

const XLSX = require('C:\\Users\\dvcha\\OneDrive\\Desktop\\Claude\\node_modules\\xlsx\\xlsx.js');
const fs   = require('fs');
const path = require('path');

// ── Files to process ──────────────────────────────────────────────────────────
const FILES = [
  "C:\\Users\\dvcha\\Voris Ventures\\Gridline Utility Group - Documents\\02. Weekly Ticket Count\\CompletionDetailReportByMemberCode (1).xls",
  "C:\\Users\\dvcha\\Voris Ventures\\Gridline Utility Group - Documents\\02. Weekly Ticket Count\\CompletionDetailReportByMemberCode (2).xls",
  "C:\\Users\\dvcha\\Voris Ventures\\Gridline Utility Group - Documents\\02. Weekly Ticket Count\\CompletionDetailReportByMemberCode (3).xls",
];

// ── Helpers ───────────────────────────────────────────────────────────────────
function excelToDate(serial) {
  if (!serial || isNaN(serial)) return null;
  return new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * 86400000)
    .toISOString().slice(0, 10);
}

// ── Parse one file ─────────────────────────────────────────────────────────
function parseExcelFile(filePath) {
  const filename = path.basename(filePath);
  console.log(`  Parsing: ${filename}`);

  const wb = XLSX.readFile(filePath);
  const sheetName = wb.SheetNames.includes('ReportData') ? 'ReportData' : wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '' });

  let fromDate = null, throughDate = null;
  for (let i = 0; i < 25; i++) {
    const r = rows[i];
    for (let j = 0; j < r.length - 1; j++) {
      const v = String(r[j]);
      if (v.includes('From Date'))    { for (let k=j+1;k<r.length;k++) if (r[k]&&!isNaN(r[k])) { fromDate    = excelToDate(r[k]); break; } }
      if (v.includes('Through Date')) { for (let k=j+1;k<r.length;k++) if (r[k]&&!isNaN(r[k])) { throughDate = excelToDate(r[k]); break; } }
    }
  }

  const tickets = [];
  let currentMember = '';

  for (let i = 0; i < rows.length; i++) {
    const r    = rows[i];
    const col1 = String(r[1] || '');

    if (col1.includes('Member Code:')) { currentMember = col1.replace('Member Code:', '').trim(); continue; }
    if (!/^[A-Z]{2}\d{8,}/.test(col1)) continue;

    const r3 = rows[i+2] || [], r5 = rows[i+4] || [];
    const rawStatus = String(r[10] || '').trim().toUpperCase();
    const status = rawStatus === 'NOT COMPLETE' ? 'incomplete'
                 : rawStatus === 'MARKED'        ? 'marked'
                 : rawStatus === 'CLEARED'        ? 'cleared' : rawStatus.toLowerCase();

    const dueDate        = excelToDate(r[15]);
    const completionDate = excelToDate(Math.floor(r[19]));
    const onTime         = dueDate && completionDate ? completionDate <= dueDate : null;

    tickets.push({
      member:        currentMember,
      ticketNum:     col1,
      locator:       String(r[3] || ''),
      operator:      String(r[6] || ''),
      status,
      dueDate,
      completionDate,
      onTime,
      workType:      String(r3[15] || '').replace('Work Type:', '').trim(),
      workPerformed: String(r5[6]  || ''),
    });
  }

  console.log(`    → ${tickets.length} tickets | ${fromDate} – ${throughDate}`);
  return { filename, fromDate, throughDate, tickets };
}

// ── Main ──────────────────────────────────────────────────────────────────────
console.log('Building data.json...\n');

const weeks = FILES
  .filter(f => { const exists = fs.existsSync(f); if (!exists) console.warn(`  SKIP (not found): ${path.basename(f)}`); return exists; })
  .map(f => parseExcelFile(f))
  .sort((a, b) => (a.fromDate || '') < (b.fromDate || '') ? -1 : 1);

const output = {
  lastUpdated: new Date().toISOString(),
  weeks
};

const outPath = path.join(__dirname, 'data.json');
fs.writeFileSync(outPath, JSON.stringify(output, null, 2));

console.log(`\n✓ Wrote data.json`);
console.log(`  Weeks : ${weeks.length}`);
console.log(`  Tickets: ${weeks.reduce((s, w) => s + w.tickets.length, 0).toLocaleString()}`);
console.log(`  Range  : ${weeks[0]?.fromDate} → ${weeks[weeks.length-1]?.throughDate}`);
