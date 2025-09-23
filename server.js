import express from "express";
import cors from "cors";
import multer from "multer";
import XLSX from "xlsx";
import morgan from "morgan";
import rateLimit from "express-rate-limit";



const app = express();

// --- базовые мидлвары ---
app.use(morgan(process.env.NODE_ENV === "production" ? "combined" : "dev"));

const allowed = (process.env.CORS_ORIGINS || "").split(",").filter(Boolean);
app.use(cors(allowed.length ? { origin: allowed } : undefined));

app.use(rateLimit({ windowMs: 15 * 60 * 1000, max: 60 })); // 60 запросов / 15 мин

// простая авторизация по токену (по желанию)
app.use((req, res, next) => {
  const expected = process.env.API_TOKEN;
  if (!expected) return next();
  if (req.header("x-api-token") === expected) return next();
  return res.status(401).json({ error: "Unauthorized" });
});

// --- загрузка файлов в память ---
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: (process.env.MAX_FILE_SIZE_MB ? +process.env.MAX_FILE_SIZE_MB : 10) * 1024 * 1024 }
});

// --- health & version ---
app.get("/healthz", (req, res) => {
  res.status(200).send("ok-v2");
});
app.get("/version", (req, res) => res.json({ version: process.env.VERSION || "dev" }));

// --- главный эндпоинт: XLS/XLSX -> XLSX ---
app.post("/convert/xlsx", upload.single("file"), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "Attach file field name 'file'." });

    // Разрешаем только .xls/.xlsx
    const name = (req.file.originalname || "").toLowerCase();
    if (!/\.(xls|xlsx)$/.test(name)) {
      return res.status(400).json({ error: "Only .xls or .xlsx files are allowed." });
    }

    // Читаем как буфер (XLS старого формата поддерживается)
    const workbook = XLSX.read(req.file.buffer, { type: "buffer" });

    // Пишем в XLSX
    const out = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    const base = (req.file.originalname || "converted").replace(/\.[^.]+$/, "");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${base}.xlsx"`);
    return res.send(out);
  } catch (e) {
    console.error(e);
    // Частая причина — битый/защищённый файл
    return res.status(500).json({ error: "Failed to convert to XLSX. Make sure the file is a valid .xls/.xlsx." });
  }
});

// === /api/statement/parse (stub) ===

const uploadXlsOnly = multer({
  limits: { fileSize: 10 * 1024 * 1024 }, // 10 MB
  fileFilter: (req, file, cb) => {
    const name = (file.originalname || "").toLowerCase();
    if (!name.endsWith(".xls")) return cb(new Error("ONLY_XLS_ALLOWED"));
    cb(null, true);
  },
});

// === helpers ===
function parseKgsNumber(val) {
  if (typeof val !== "string" && typeof val !== "number") return NaN;
  let s = String(val).trim();
  if (!s) return NaN;
  s = s.replace(/\s/g, "");
  const commaAsDecimal = /,\d{1,2}$/.test(s);
  if (commaAsDecimal) s = s.replace(/\./g, "").replace(",", ".");
  return Number(s);
}
function tryParseDate(d) {
  if (!d) return null;
  if (d instanceof Date && !isNaN(d)) return d;
  const s = String(d).trim();

  const n = Number(s); // Excel serial?
  if (!isNaN(n) && n > 25569 && n < 60000) {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(epoch.getTime() + n * 86400000);
  }
  const m = s.match(/^(\d{2})\.(\d{2})\.(\d{4})/); // dd.mm.yyyy
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})/); // yyyy-mm-dd
  if (m2) return new Date(Number(m2[1]), Number(m2[2]) - 1, Number(m2[3]));

  const dt = new Date(s);
  return isNaN(dt) ? null : dt;
}
function ymd(dt) {
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, "0");
  const d = String(dt.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}
const COLS = {
  date: ["Дата", "Дата операции", "Operation date", "Posting date", "Дата проводки"],
  desc: ["Описание", "Описание операции", "Description", "Назначение платежа", "Назначение"],
  debit: ["Списание", "Дебет", "Расход", "Debit"],
  credit: ["Поступление", "Кредит", "Доход", "Credit"],
  amount: ["Сумма", "Amount", "Итого"],
};
function pick(row, keys) {
  const map = Object.fromEntries(Object.keys(row).map(k => [k.toLowerCase().trim(), k]));
  for (const k of keys) {
    const hit = map[k.toLowerCase()];
    if (hit) return row[hit];
  }
  return undefined;
}
// === end helpers ===


app.post("/api/statement/parse", uploadXlsOnly.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "NO_FILE" });

    const wb = XLSX.read(req.file.buffer, { type: "buffer" });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];

    const isDebug = String(req.query?.debug || "") === "1";
    const HEADER_INDEX = 12; // 13-я строка

    // читаем, считая 13-ю строку заголовками
    const rowsFixed = XLSX.utils.sheet_to_json(ws, { defval: "", range: HEADER_INDEX });

    if (isDebug) {
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      return res.json({
        firstSheet: sheetName,
        headerIndexUsed: HEADER_INDEX,
        headerRowUsed: aoa[HEADER_INDEX] || [],
        rowsLenFixed: rowsFixed.length,
        sampleRowFixed: rowsFixed[0] || null,
        previewFixed: rowsFixed.slice(0, 3),
      });
    }

    // -------- боевой парсинг --------
    const transactions = [];
    for (const r of rowsFixed) {
      const dateRaw = pick(r, COLS.date);
      const dt = tryParseDate(dateRaw);
      if (!dt) continue; // игнорируем хвост/строки без даты

      const desc = (pick(r, COLS.desc) ?? "").toString().trim();

      const debitRaw = pick(r, COLS.debit);
      const creditRaw = pick(r, COLS.credit);
      const amountRaw = pick(r, COLS.amount);

      let spending = NaN;
      if (debitRaw !== undefined && String(debitRaw).trim() !== "") {
        spending = parseKgsNumber(debitRaw);
      } else if (amountRaw !== undefined && String(amountRaw).trim() !== "") {
        const a = parseKgsNumber(amountRaw);
        const isIncome = creditRaw !== undefined && String(creditRaw).trim() !== "" && parseKgsNumber(creditRaw) > 0;
        spending = isIncome ? 0 : (a < 0 ? Math.abs(a) : a);
      }
      if (isNaN(spending) || spending <= 0) continue;

      transactions.push({
        date: ymd(dt),
        description: desc,
        amount: -Math.abs(spending), // траты — отрицательные
      });
    }

    let from = null, to = null;
    if (transactions.length) {
      const dates = transactions.map(t => t.date).sort();
      from = dates[0];
      to = dates[dates.length - 1];
    }

    const byDay = new Map();
    for (const t of transactions) {
      byDay.set(t.date, (byDay.get(t.date) || 0) + Math.abs(t.amount));
    }
    const dailySpending = Array.from(byDay.entries())
      .sort((a, b) => (a[0] < b[0] ? -1 : 1))
      .map(([date, amount]) => ({ date, amount: Number(amount.toFixed(2)) }));

    const totals = {
      spending: Number(transactions.reduce((acc, t) => acc + Math.abs(t.amount), 0).toFixed(2)),
    };

    return res.json({ period: { from, to }, dailySpending, transactions, totals });
    // -------- /боевой парсинг --------
  } catch (e) {
    if (e?.message === "ONLY_XLS_ALLOWED") {
      return res.status(415).json({ error: "ONLY_XLS_ALLOWED" });
    }
    console.error(e);
    return res.status(500).json({ error: "PARSE_FAILED" });
  }
});



// 404
app.use((_req, res) => res.status(404).json({ error: "Not found" }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`XLS->XLSX converter listening on ${PORT}`));
