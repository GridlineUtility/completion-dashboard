const XLSX = require('xlsx');
const fs   = require('fs');
const path = require('path');

const BASE = 'C:\\Users\\dvcha\\Gridline Utility Group\\Gridline Utility Group - Documents\\02. Productivity Dashboards\\01. Weekly Ticket Count\\';
const FILES = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19].map(n => BASE + `CompletionDetailReportByMemberCode (${n}).xls`);

function excelToDate(serial) {
  if (!serial || isNaN(serial)) return null;
  return new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * 86400000).toISOString().slice(0, 10);
}

function parseFile(filePath) {
  const filename = path.basename(filePath);
  process.stdout.write(`Parsing: ${filename} ... `);
  const wb = XLSX.readFile(filePath);
  const sheetName = wb.SheetNames.includes('ReportData') ? 'ReportData' : wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '' });

  let fromDate = null, throughDate = null;
  for (let i = 0; i < 25; i++) {
    const r = rows[i] || [];
    for (let j = 0; j < r.length - 1; j++) {
      const v = String(r[j]);
      if (v.includes('From Date'))    { for (let k=j+1;k<r.length;k++) if (r[k]&&!isNaN(r[k])) { fromDate    = excelToDate(r[k]); break; } }
      if (v.includes('Through Date')) { for (let k=j+1;k<r.length;k++) if (r[k]&&!isNaN(r[k])) { throughDate = excelToDate(r[k]); break; } }
    }
  }

  const tickets = [];
  let currentMember = '';
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || [], col1 = String(r[1] || '');
    if (col1.includes('Member Code:')) { currentMember = col1.replace('Member Code:', '').trim(); continue; }
    if (!/^[A-Z]{2}\d{8,}/.test(col1)) continue;
    const r3 = rows[i+2] || [], r5 = rows[i+4] || [];
    const rawStatus = String(r[10] || '').trim().toUpperCase();
    const status = rawStatus === 'NOT COMPLETE' ? 'incomplete' : rawStatus === 'MARKED' ? 'marked' : rawStatus === 'CLEARED' ? 'cleared' : rawStatus.toLowerCase();
    const dueDate = excelToDate(r[15]), completionDate = excelToDate(Math.floor(r[19]));
    const onTime = dueDate && completionDate ? completionDate <= dueDate : null;
    tickets.push({ member: currentMember, ticketNum: col1, locator: String(r[3]||''), operator: String(r[6]||''), status, dueDate, completionDate, onTime, workType: String(r3[15]||'').replace('Work Type:','').trim(), workPerformed: String(r5[6]||'') });
  }
  console.log(`${fromDate} to ${throughDate}: ${tickets.length} tickets`);
  return { filename, fromDate, throughDate, tickets };
}

const weeks = FILES.map(parseFile).sort((a,b) => (a.fromDate||'') < (b.fromDate||'') ? -1 : 1);
const output = { lastUpdated: new Date().toISOString(), weeks };
fs.writeFileSync(path.join(__dirname, 'data.json'), JSON.stringify(output));
console.log(`\nWrote data.json — ${weeks.length} weeks, ${weeks.reduce((s,w)=>s+w.tickets.length,0).toLocaleString()} total tickets`);



