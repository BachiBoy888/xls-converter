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
app.get("/healthz", (req, res) => res.status(200).send("ok"));
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

app.post("/api/statement/parse", uploadXlsOnly.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "NO_FILE" });

    // Проверим, что файл читается как Excel
    XLSX.read(req.file.buffer, { type: "buffer" });

    // Возвращаем минимальный контракт-заглушку
      return res.json({
            period: { from: null, to: null },
            dailySpending: [],
            transactions: [],
            totals: { spending: 0 },
          });
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
