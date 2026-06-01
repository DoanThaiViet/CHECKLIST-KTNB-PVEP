const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Nguồn: file Excel đồng bộ từ M365/SharePoint về máy (OneDrive).
const SRC = process.env.CHECKLIST_XLSX ||
  'C:\\Users\\vietdt\\OneDrive - PetroVietnam Exploration Production Corporation (PVEP)\\M365 - Kiểm soát nội bộ - Documents\\6.2_KTNB 2026\\BMĐH\\Phân công nhiệm vụ và Checklist.xlsx';

// Đích: checklists.json ở gốc repo (đúng nơi web app đọc).
const DEST = path.join(__dirname, 'checklists.json');

const wb = XLSX.readFile(SRC);
const result = {};

wb.SheetNames.forEach(sheetName => {
  const ws = wb.Sheets[sheetName];
  if (!ws || !ws['!ref']) return;
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  let headerIdx = -1;
  for (let i = 0; i < Math.min(8, rows.length); i++) {
    if (rows[i] && rows[i][0] === 'STT') { headerIdx = i; break; }
  }
  if (headerIdx === -1) return;

  const headers = rows[headerIdx].filter(h => h !== '');
  const groups = [];
  let currentGroup = { name: '', items: [] };

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.every(c => c === '' || c === undefined || c === null)) continue;
    const col0 = row[0], col2 = row[2] || '';
    const isGroupHeader = (typeof col0 === 'string' && col0.length > 3 && isNaN(Number(col0)) && col0 !== 'STT' && col2 === '');

    if (isGroupHeader) {
      if (currentGroup.items.length > 0 || currentGroup.name) groups.push(currentGroup);
      currentGroup = { name: col0.trim(), items: [] };
      continue;
    }

    const item = {};
    headers.forEach((h, idx) => {
      if (h && row[idx] !== undefined && row[idx] !== '') item[h] = String(row[idx]).trim();
    });
    if (Object.keys(item).length > 1) currentGroup.items.push(item);
  }
  if (currentGroup.items.length > 0) groups.push(currentGroup);

  const totalItems = groups.reduce((s, g) => s + g.items.length, 0);
  if (totalItems < 1) return;
  const hStr = headers.join('|').toLowerCase();
  if (hStr.includes('bước') || hStr.includes('kiểm tra') || hStr.includes('mục tiêu') || hStr.includes('nội dung') || hStr.includes('hệ thống') || hStr.includes('hạng mục')) {
    result[sheetName] = { headers, groups };
  }
});

fs.writeFileSync(DEST, JSON.stringify(result, null, 2), 'utf8');

const totalSheets = Object.keys(result).length;
let totalGroups = 0, totalItems = 0;
Object.values(result).forEach(d => d.groups.forEach(g => { totalGroups++; totalItems += g.items.length; }));
console.log(`SUCCESS: Sheets=${totalSheets}, Groups=${totalGroups}, Items=${totalItems}`);
