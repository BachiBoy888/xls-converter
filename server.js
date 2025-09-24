import express from 'express';
import cors from 'cors';
import multer from 'multer';
import XLSX from 'xlsx';
import morgan from 'morgan';
import rateLimit from 'express-rate-limit';
import { DateTime } from 'luxon';
import { v4 as uuidv4 } from 'uuid';
import fs from 'fs/promises';
import os from 'os';
import helmet from 'helmet';


const app = express();
app.use(helmet({ crossOriginResourcePolicy: false }));

// CORS: либо из ENV, либо '*'
const allowed = (process.env.CORS_ORIGINS || '').split(',').map(s => s.trim()).filter(Boolean);
app.use(cors(allowed.length ? {
  origin: allowed,
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'x-api-token'],
  maxAge: 600
} : {
  origin: '*',
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'x-api-token'],
  maxAge: 600
}));
app.options('*', cors());

app.set('trust proxy', 1); // чтобы rateLimit видел реальный IP


// === requestId + таймер (ставим ДО morgan/rateLimit) ===
app.use((req, res, next) => {
  req.id = uuidv4();
  req._startMs = Date.now();
  next();
});

// morgan с requestId
morgan.token('rid', req => req.id);
app.use(morgan((process.env.NODE_ENV === 'production' ? 'combined' : 'dev') + ' :rid'));

app.use(rateLimit({ windowMs: 15 * 60 * 1000, max: 60 })); // 60 запросов / 15 мин

// === опциональная авторизация по токену ===
app.use((req, res, next) => {
  const expected = process.env.API_TOKEN;
  if (!expected) return next();
  if (req.header('x-api-token') === expected) return next();
  return res.status(401).json({ error: 'Unauthorized', requestId: req.id });
});

// --- загрузка файлов в память ---
const upload = multer({
    storage: multer.memoryStorage(),
      limits: { fileSize: 20 * 1024 * 1024 } // 20 MB
});

// --- health & version ---
app.get("/healthz", (req, res) => {
  res.status(200).send("ok-v2");
});
app.get("/version", (req, res) => res.json({ version: process.env.VERSION || "dev" }));

// --- главный эндпоинт: XLS/XLSX -> XLSX ---
app.post("/convert/xlsx", upload.single("file"), (req, res) => {
    const rid = req.id;
      const started = Date.now();
      try {
        if (!req.file) return res.status(400).json({ error: "NO_FILE", requestId: rid });

        const name = (req.file.originalname || "").toLowerCase();
        if (!/\.(xls|xlsx)$/.test(name)) {
          return res.status(400).json({ error: "ONLY_XLS_OR_XLSX", requestId: rid });
        }

        const workbook = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
        const out = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

        const base = (req.file.originalname || "converted").replace(/\.[^.]+$/, "");
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-Disposition", `attachment; filename="${base}.xlsx"`);

        const parseMs = Date.now() - started;
        console.log(`[convert] rid=${rid} file="${req.file.originalname}" size=${req.file.size}B parseMs=${parseMs}`);
        return res.send(out);
      } catch (e) {
        console.error(`[convert][error] rid=${rid}`, e);
        return res.status(500).json({ error: "CONVERT_FAILED", requestId: rid });
      }
});

// === upload для /api/statement/parse: сохраняем во временный файл в /tmp
const uploadXlsOnly = multer({
  storage: multer.diskStorage({
    destination: (_req, _file, cb) => cb(null, os.tmpdir()),
    filename: (_req, file, cb) => {
      const base = (file.originalname || 'upload').replace(/\s+/g, '_');
      cb(null, `${Date.now()}_${base}`);
    }
  }),
  limits: { fileSize: 20 * 1024 * 1024 }, // 20 MB
  fileFilter: (_req, file, cb) => {
    const name = (file.originalname || '').toLowerCase();
    if (!name.endsWith('.xls')) return cb(new Error('ONLY_XLS_ALLOWED'));
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

  // Excel serial?
  const n = Number(s);
  if (!isNaN(n) && n > 25569 && n < 60000) {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(epoch.getTime() + n * 86400000);
  }

  // dd.mm.yyyy
  const m = s.match(/^(\d{2})\.(\d{2})\.(\d{4})/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));

  // yyyy-mm-dd
  const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
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
  date: ["Дата", "Дата операции", "Operation date", "Posting date", "Дата проводки", "Date"],
  desc: [
    "Operation", "Recipient/Payer",       // MBank столбцы
    "Описание", "Описание операции", "Description", "Назначение платежа", "Назначение"
  ],
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



app.post('/api/statement/parse', uploadXlsOnly.single('file'), async (req, res) => {
    const rid = req.id;
      const started = Date.now();

      // фиксированная таймзона
      const zone = 'Asia/Bishkek';

      try {
        if (!req.file) return res.status(400).json({ error: 'NO_FILE' });

        const filepath = req.file.path;
        const originalName = req.file.originalname;
        const size = req.file.size;

        // читаем файл с диска (xls/xlsx поддерживается)
        const buf = await fs.readFile(filepath);
          const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });
        const sheetName = wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];

        const isDebug = String(req.query?.debug || '') === '1';
        const HEADER_INDEX = 12; // 13-я строка — как ты и ставил

        const rowsFixed = XLSX.utils.sheet_to_json(ws, { defval: '', range: HEADER_INDEX });

        if (isDebug) {
          const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          const parseMs = Date.now() - started;
          return res.json({
            meta: {
              requestId: rid,
              processedAt: DateTime.now().setZone(zone).toISO(),
              parseMs
            },
            debug: {
              firstSheet: sheetName,
              headerIndexUsed: HEADER_INDEX,
              headerRowUsed: aoa[HEADER_INDEX] || [],
              rowsLenFixed: rowsFixed.length,
              sampleRowFixed: rowsFixed[0] || null,
              previewFixed: rowsFixed.slice(0, 3),
            }
          });
        }

        // ---- парсинг транзакций (твоя логика с небольшими правками) ----
          const transactions = [];
          for (const r of rowsFixed) {
            const dateRaw = pick(r, COLS.date);
            const dt = tryParseDate(dateRaw);
            if (!dt) continue;

            const desc = (pick(r, COLS.desc) ?? '').toString().trim();
              // ВНИМАНИЕ: в этом файле:
              // creditRaw -> РАСХОД (списание)
              // debitRaw  -> ПРИХОД (зачисление)
              const creditRaw = pick(r, COLS.credit); // расход
              const debitRaw  = pick(r, COLS.debit);  // приход

              const expense = Number(parseKgsNumber(creditRaw)) || 0; // списание (>=0)
              const income  = Number(parseKgsNumber(debitRaw))  || 0; // зачисление (>=0)

              // если нет ни прихода, ни расхода — пропускаем строку
                if (expense <= 0 && income <= 0) continue;

                // amount по контракту: приход (+), расход (-)
                const amount = income - expense;

              transactions.push({
                 date: ymd(dt),           // YYYY-MM-DD
                 description: desc,
                 amount,                  // приход > 0, расход < 0
                 credit: income,          // зачисления (плюсом)
                 debit:  expense,         // списания (плюсом)
               });
             }

          // период
          let from = null, to = null;
          if (transactions.length) {
            const dates = transactions.map(t => t.date).sort();
            from = dates[0];
            to   = dates[dates.length - 1];
          }
          
          // агрегаты по дням
          const byDay = new Map();
          /*
            Накапливаем по дням:
            - credit: сумма зачислений
            - debit : сумма списаний
          */
          for (const t of transactions) {
            const cur = byDay.get(t.date) || { date: t.date, credit: 0, debit: 0 };
            cur.credit += t.credit || 0;
            cur.debit  += t.debit  || 0;
            byDay.set(t.date, cur);
          }
          
          // заполнение «пустых дней» нулями на периоде
          const dailySpending = [];
          if (from && to) {
            let cur = DateTime.fromISO(from, { zone });
            const end = DateTime.fromISO(to, { zone });
            while (cur <= end) {
              const key = cur.toISODate(); // YYYY-MM-DD
              const v = byDay.get(key) || { date: key, credit: 0, debit: 0 };
              dailySpending.push({
                date: v.date,
                credit: Number(v.credit.toFixed(2)),
                debit : Number(v.debit.toFixed(2)),
                net   : Number((v.credit - v.debit).toFixed(2)),
                // сохранить старое поле amount как «расходы по модулю» (для графика)
                amount: Number(v.debit.toFixed(2)),
              });
              cur = cur.plus({ days: 1 });
            }
          }
          

          // totals
          const totals = (() => {
            const credits = transactions.reduce((s, t) => s + (t.credit || 0), 0);
            const debits  = transactions.reduce((s, t) => s + (t.debit  || 0), 0);
            const net     = credits - debits;
            return {
              credits: Number(credits.toFixed(2)),  // сумма зачислений
              debits : Number(debits.toFixed(2)),   // сумма списаний
              net    : Number(net.toFixed(2)),      // нетто
              expenses: Number(debits.toFixed(2)),  // совместимость со старым полем
            };
          })();

        const parseMs = Date.now() - started;

        // account.* (как просили)
        const account = {
          currency: 'KGS',
          bank: 'MBank'
        };

        // meta.*
        const meta = {
          processedAt: DateTime.now().setZone(zone).toISO(), // ISO с +06:00
          requestId: rid,
          file: {
            name: originalName,
            size
          },
          sheet: sheetName,
          rows: rowsFixed.length,
          parseMs
        };

        // логи (консоль)
        console.log(`[parse] rid=${rid} file="${originalName}" size=${size}B rows=${rowsFixed.length} parseMs=${parseMs}`);

        return res.json({ meta, account, period: { from, to }, dailySpending, transactions, totals });

      } catch (e) {
        if (e?.message === 'ONLY_XLS_ALLOWED') {
          return res.status(415).json({ error: 'ONLY_XLS_ALLOWED' });
        }
        console.error(`[parse][error] rid=${rid}`, e);
        return res.status(500).json({ error: 'PARSE_FAILED', requestId: rid });
      } finally {
        // удалить временный файл
        if (req?.file?.path) {
          try { await fs.unlink(req.file.path); } catch {}
        }
      }
});

// --- error handler для multer и наших ошибок ---
app.use((err, req, res, next) => {
    if (err?.message === 'ONLY_XLS_ALLOWED') {
      return res.status(415).json({ error: 'ONLY_XLS_ALLOWED' });
    }
    if (err && err instanceof multer.MulterError) {
      if (err.code === 'LIMIT_FILE_SIZE') {
        return res.status(413).json({ error: 'PAYLOAD_TOO_LARGE' }); // по требованию
      }
      return res.status(400).json({ error: 'UPLOAD_ERROR', code: err.code });
    }
    return next(err);
});

// 404
app.use((_req, res) => res.status(404).json({ error: "Not found" }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`XLS->XLSX converter listening on ${PORT}`));


