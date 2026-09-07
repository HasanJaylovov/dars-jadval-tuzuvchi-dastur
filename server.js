import express from "express";
import cors from "cors";
import multer from "multer";
import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";
import crypto from "crypto";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const app = express();
const PORT = process.env.PORT || 3000;
const uploadDir = path.join(__dirname, "uploads");
fs.mkdirSync(uploadDir, { recursive: true });

app.use(cors({ origin: "*" }));
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, "public")));

const storage = multer.diskStorage({
  destination: uploadDir,
  filename: (req, file, cb) => cb(null, `jadval-${Date.now()}-${crypto.randomBytes(4).toString("hex")}.xlsx`)
});
const upload = multer({
  storage,
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === ".xlsx") cb(null, true);
    else cb(new Error("Faqat .xlsx formatidagi Excel fayl qabul qilinadi."));
  }
});

const sessions = new Map();

function normalizeCellValue(value) {
  if (value === undefined || value === null) return "";
  if (value instanceof Date) return value.toISOString();
  if (typeof value === "object") {
    if (Array.isArray(value.richText)) return value.richText.map(x => x?.text || "").join("");
    if (Object.prototype.hasOwnProperty.call(value, "result")) return normalizeCellValue(value.result);
    if (value.text !== undefined) return String(value.text);
    if (value.hyperlink) return String(value.hyperlink);
    if (value.error) return String(value.error);
    return "";
  }
  return String(value);
}

async function readExcel(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  return {
    sheets: workbook.worksheets.map(ws => {
      const rows = [];
      const max = Math.max(ws.columnCount || 0, 1);
      for (let r = 1; r <= (ws.rowCount || 1); r++) {
        const row = [];
        for (let c = 1; c <= max; c++) row.push(normalizeCellValue(ws.getCell(r, c).value));
        rows.push(row);
      }
      return { name: ws.name, rowCount: ws.rowCount || rows.length, columnCount: max, rows };
    })
  };
}

function looksLikeTeacherId(value) {
  const s = String(value ?? "").trim();
  return /^\d{1,4}$/.test(s);
}
function clean(v) { return String(v ?? "").trim(); }
function sameTeacher(a, b) { return clean(a) !== "" && clean(a) === clean(b) && looksLikeTeacherId(a); }

// Generic parser for the real workbook: schedule blocks normally use
// [teacher ID, subject] pairs, with class name one row above the pair.
function detectConflicts(sheet) {
  const rows = sheet.rows || [];
  const conflicts = [];
  const slots = new Map();

  // Find teacher-ID cells and infer their class from nearby header rows.
  for (let r = 0; r < rows.length; r++) {
    const period = clean(rows[r]?.[1]);
    if (!/^\d{1,2}$/.test(period)) continue;
    const day = clean(rows[r]?.[0]);
    if (!day) continue;

    for (let c = 0; c < (rows[r]?.length || 0); c++) {
      const teacher = clean(rows[r]?.[c]);
      if (!looksLikeTeacherId(teacher)) continue;
      const subject = clean(rows[r]?.[c + 1]);
      if (!subject) continue;

      let className = "";
      for (let rr = Math.max(0, r - 3); rr < r; rr++) {
        const candidate = clean(rows[rr]?.[c]);
        if (candidate && /sinf/i.test(candidate)) className = candidate;
      }
      // Most sheets have class headers two rows above; merged headers can put it in c-1.
      if (!className) {
        for (let rr = Math.max(0, r - 3); rr < r; rr++) {
          for (const cc of [c, c - 1, c + 1]) {
            const candidate = clean(rows[rr]?.[cc]);
            if (candidate && /sinf/i.test(candidate)) { className = candidate; break; }
          }
          if (className) break;
        }
      }
      const key = `${day.toUpperCase()}|${period}|${teacher}`;
      const item = { row: r, col: c, teacherId: teacher, subject, className: className || `ustun ${c + 1}` };
      if (!slots.has(key)) slots.set(key, []);
      slots.get(key).push(item);
    }
  }

  for (const [key, items] of slots) {
    if (items.length > 1) {
      conflicts.push({
        id: crypto.createHash("md5").update(sheet.name + key).digest("hex").slice(0, 10),
        type: "teacher",
        day: key.split("|")[0],
        period: key.split("|")[1],
        teacherId: key.split("|")[2],
        items
      });
    }
  }
  return conflicts;
}

function findFreeSlots(sheet, teacherId, sourceCol = null) {
  const rows = sheet.rows || [];
  const busyTeacher = new Set();
  const candidates = [];

  // A teacher is busy if their ID occurs in the same day/period row.
  for (let r = 0; r < rows.length; r++) {
    const period = clean(rows[r]?.[1]);
    const day = clean(rows[r]?.[0]);
    if (!/^\d{1,2}$/.test(period) || !day) continue;
    for (let c = 0; c < (rows[r]?.length || 0); c++) {
      const teacher = clean(rows[r]?.[c]);
      if (looksLikeTeacherId(teacher)) busyTeacher.add(`${day}|${period}|${teacher}`);
    }
  }

  // The destination must be a genuinely empty lesson cell in the SAME class column
  // (teacher ID + subject pair). Never overwrite another lesson.
  for (let r = 0; r < rows.length; r++) {
    const period = clean(rows[r]?.[1]);
    const day = clean(rows[r]?.[0]);
    if (!/^\d{1,2}$/.test(period) || !day) continue;
    if (busyTeacher.has(`${day}|${period}|${teacherId}`)) continue;

    if (sourceCol !== null && Number.isInteger(sourceCol)) {
      const teacherCell = clean(rows[r]?.[sourceCol]);
      const subjectCell = clean(rows[r]?.[sourceCol + 1]);
      if (teacherCell !== '' || subjectCell !== '') continue;
      candidates.push({ day, period, row: r, col: sourceCol });
    } else {
      candidates.push({ day, period, row: r });
    }
  }
  return candidates.slice(0, 30);
}

app.get("/api/health", (req, res) => res.json({ ok: true, message: "Dars Jadvali server ishlayapti" }));

app.post("/api/upload", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ ok: false, error: "Fayl tanlanmagan." });
    const data = await readExcel(req.file.path);
    const token = crypto.randomBytes(16).toString("hex");
    sessions.set(token, { path: req.file.path, originalName: req.file.originalname, created: Date.now() });
    res.json({ ok: true, token, file: { originalName: req.file.originalname, filename: req.file.filename, size: req.file.size }, ...data });
  } catch (e) {
    if (req.file?.path) try { fs.unlinkSync(req.file.path); } catch {}
    res.status(500).json({ ok: false, error: "Excel faylni o'qishda xatolik.", details: e.message });
  }
});

app.post("/api/check", (req, res) => {
  try {
    const sheets = Array.isArray(req.body?.sheets) ? req.body.sheets : req.body?.sheet ? [req.body.sheet] : [];
    const all = [];
    for (const sheet of sheets) {
      for (const c of detectConflicts(sheet)) all.push({ ...c, sheet: sheet.name });
    }
    res.json({ ok: true, count: all.length, conflicts: all });
  } catch (e) { res.status(500).json({ ok: false, error: "Tekshirish xatosi", details: e.message }); }
});

app.post("/api/suggestions", (req, res) => {
  const { sheet, teacherId } = req.body || {};
  if (!sheet || !teacherId) return res.status(400).json({ ok: false, error: "Sheet va Teacher ID kerak." });
  res.json({ ok: true, suggestions: findFreeSlots(sheet, String(teacherId), Number.isInteger(req.body?.sourceCol) ? req.body.sourceCol : null) });
});

app.post("/api/export", async (req, res) => {
  try {
    const { token, sheets } = req.body || {};
    const session = sessions.get(token);
    if (!session || !fs.existsSync(session.path)) return res.status(400).json({ ok: false, error: "Excel sessiyasi topilmadi. Faylni qayta yuklang." });
    if (!Array.isArray(sheets)) return res.status(400).json({ ok: false, error: "Excel ma'lumotlari yuborilmadi." });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(session.path);
    for (const payload of sheets) {
      const ws = workbook.getWorksheet(payload.name);
      if (!ws || !Array.isArray(payload.rows)) continue;
      for (let r = 0; r < payload.rows.length; r++) {
        const row = payload.rows[r] || [];
        for (let c = 0; c < row.length; c++) ws.getCell(r + 1, c + 1).value = row[c] === "" ? null : row[c];
      }
    }
    const out = path.join(uploadDir, `dars-jadvali-${Date.now()}.xlsx`);
    await workbook.xlsx.writeFile(out);
    res.download(out, "dars_jadvali_natija.xlsx", () => { try { fs.unlinkSync(out); } catch {} });
  } catch (e) { res.status(500).json({ ok: false, error: "Excel eksportida xatolik.", details: e.message }); }
});

app.use((req, res) => res.status(404).json({ error: "Endpoint topilmadi.", path: req.path }));
app.use((err, req, res, next) => { console.error(err); res.status(500).json({ ok: false, error: err.message || "Server xatosi" }); });

app.listen(PORT, "0.0.0.0", () => console.log(`🚀 Dars Jadvali server running on port ${PORT}`));
