const express = require('express');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Ensure data directory
const DATA_DIR = path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

// ===== Helper: read/write JSON =====
function readJSON(file, fallback = {}) {
  const p = path.join(DATA_DIR, file);
  if (!fs.existsSync(p)) return fallback;
  try { return JSON.parse(fs.readFileSync(p, 'utf8')); }
  catch { return fallback; }
}
function writeJSON(file, data) {
  fs.writeFileSync(path.join(DATA_DIR, file), JSON.stringify(data, null, 2), 'utf8');
}

// ===== STAFF LIST =====
const STAFF = [
  { id: 1, name: "Đoàn Thái Việt", role: "leader" },
  { id: 2, name: "Đặng Ngọc Khánh", role: "leader" },
  { id: 3, name: "Vũ Anh Quân", role: "leader" },
  { id: 4, name: "Nguyễn Thị Cẩm Tú", role: "leader" },
  { id: 5, name: "Nguyễn Thị Hải Yến", role: "leader" },
  { id: 6, name: "Chu Nhật Anh", role: "pic" },
  { id: 7, name: "Ngô Thị Tuấn Anh", role: "pic" },
  { id: 8, name: "Vũ Ái Ngọc Bình", role: "pic" },
  { id: 9, name: "Nguyễn Văn Điều", role: "pic" },
  { id: 10, name: "Vũ Thúy Quỳnh", role: "pic" },
  { id: 11, name: "Trần Bảo Châu", role: "pic" },
  { id: 12, name: "Lê Thùy Dương", role: "pic" },
  { id: 13, name: "Trần Ngọc Ánh", role: "pic" },
  { id: 14, name: "Nguyễn Thị Cẩm Hà", role: "pic" },
  { id: 15, name: "Lý Quốc Đạt", role: "pic" },
  { id: 16, name: "Nguyễn Bá Lập", role: "pic" },
  { id: 17, name: "Trần Phương Anh", role: "pic" },
  { id: 18, name: "Nguyễn Bá Minh", role: "pic" },
  { id: 19, name: "Giang Anh Dũng", role: "pic" },
  { id: 20, name: "Admin", role: "admin" },
];

// ===== AUTH =====
const SHARED_CREDENTIALS = { username: "KSNB", password: "pvep2026" };

app.post('/api/login', (req, res) => {
  const { username, password } = req.body;
  if (username === SHARED_CREDENTIALS.username && password === SHARED_CREDENTIALS.password) {
    return res.json({ ok: true, staff: STAFF });
  }
  res.status(401).json({ ok: false, error: "Sai tài khoản hoặc mật khẩu" });
});

app.get('/api/staff', (req, res) => {
  res.json(STAFF);
});

// ===== CHECKLIST RESULTS =====
app.get('/api/results', (req, res) => {
  res.json(readJSON('checklist-results.json', {}));
});

app.put('/api/results/:key', (req, res) => {
  const { key } = req.params;
  const { ketQua, kienNghi, updatedBy } = req.body;
  const results = readJSON('checklist-results.json', {});
  results[key] = {
    ketQua: ketQua || '',
    kienNghi: kienNghi || '',
    updatedBy: updatedBy || '',
    updatedAt: new Date().toISOString()
  };
  writeJSON('checklist-results.json', results);
  res.json({ ok: true });
});

// ===== COMMENTS =====
app.get('/api/comments', (req, res) => {
  res.json(readJSON('comments.json', []));
});

app.post('/api/comments', (req, res) => {
  const { sheet, stepSTT, user, role, text } = req.body;
  if (!text || !text.trim()) return res.status(400).json({ error: "Comment trống" });
  const comments = readJSON('comments.json', []);
  const newComment = {
    id: Date.now(),
    sheet, stepSTT,
    user: user || 'Ẩn danh',
    role: role || 'pic',
    text: text.trim(),
    timestamp: new Date().toISOString()
  };
  comments.push(newComment);
  writeJSON('comments.json', comments);
  res.json(newComment);
});

app.delete('/api/comments/:id', (req, res) => {
  const id = parseInt(req.params.id);
  let comments = readJSON('comments.json', []);
  comments = comments.filter(c => c.id !== id);
  writeJSON('comments.json', comments);
  res.json({ ok: true });
});

// ===== EXCEL UPLOAD =====
const upload = multer({ dest: path.join(__dirname, 'uploads/') });

app.post('/api/upload-excel', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file" });

  try {
    const wb = XLSX.readFile(req.file.path);
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

    // Save to public/checklists.json
    fs.writeFileSync(path.join(__dirname, 'public', 'checklists.json'), JSON.stringify(result, null, 2), 'utf8');

    // Cleanup uploaded file
    fs.unlinkSync(req.file.path);

    const totalSheets = Object.keys(result).length;
    let totalGroups = 0, totalItems = 0;
    Object.values(result).forEach(d => d.groups.forEach(g => { totalGroups++; totalItems += g.items.length; }));

    res.json({ ok: true, sheets: totalSheets, groups: totalGroups, items: totalItems });
  } catch (err) {
    if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    res.status(500).json({ error: err.message });
  }
});

// ===== START =====
app.listen(PORT, () => {
  console.log(`🟢 PVEP Dashboard running at http://localhost:${PORT}`);
});
