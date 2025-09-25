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
import { buildCumulativeTimeline, attachDailyCumulativeClose } from './utils/cumulative.js';

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
app.get('/healthz', (_req, res) => res.status(200).send('ok-v2'));
app.get('/version', (_req, res) => {
  res.json({
    name: 'xls-converter',
    version: process.env.VERSION || 'dev',
    commit: process.env.COMMIT_SHA || null,
    uptimeSec: Math.round(process.uptime())
  });
});

// --- вспомогательные функции ---
function parseKgsNumber(val) {
  if (typeof val !== 'string' && typeof val !== 'number') return NaN;
  let s = String(val).trim();
  if (!s) return NaN;
  s = s.replace(/\s/g, '');
  const commaAsDecimal = /,\d{1,2}$/.test(s);
  if (commaAsDecimal) s = s.replace(/\./g, '').replace(',', '.');
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
  const m = String(dt.getMonth() + 1).padStart(2, '0');
  const d = String(dt.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

// Вынесли ключи столбцов в константы (для унификации по разным банкам)
const COLS = {
  date: ['Дата', 'Дата операции', 'Operation date', 'Posting date', 'Дата проводки', 'Date'],
  desc: [
    'Operation', 'Recipient/Payer', // MBank
    'Описание', 'Описание операции', 'Description', 'Назначение платежа', 'Назначение'
  ],
  // ВАЖНО для MBank: по debug видно, что в выгрузке
  //   Credit -> СПИСАНИЕ (расход)
  //   Debit  -> ЗАЧИСЛЕНИЕ (приход)
  debitIsIncomeKeys: ['Debit', 'Поступление', 'Кредит', 'Доход'], // приход
  creditIsExpenseKeys: ['Credit', 'Списание', 'Дебет', 'Расход'],   // расход
  amount: ['Сумма', 'Amount', 'Итого']
};

COLS.time = ['Время', 'Time', 'Время операции', 'Operation time', 'Transaction time'];

function pick(row, keys) {
  const map = Object.fromEntries(Object.keys(row).map(k => [k.toLowerCase().trim(), k]));
  for (const k of keys) {
    const hit = map[k.toLowerCase()];
    if (hit) return row[hit];
  }
  return undefined;
}
function buildTsISO(dateRaw, timeRaw, index, zone) {
  // Если пришёл Date — используем его как есть (может уже содержать время)
  if (dateRaw instanceof Date && !isNaN(dateRaw)) {
    const dt = DateTime.fromJSDate(dateRaw, { zone });
    // если в исходном Date время=00:00:00 и есть колонка timeRaw — добавим её
    if (
      (dt.hour + dt.minute + dt.second === 0) &&
      timeRaw !== undefined && timeRaw !== null && String(timeRaw).trim() !== ''
    ) {
      const t = normalizeTimeFromAny(timeRaw);
      if (t) {
        const composed = DateTime.fromISO(
          `${dt.toFormat('yyyy-LL-dd')}T${t}`,
          { zone }
        );
        return composed.isValid ? composed.toISO() : null;
      }
    }
    return dt.isValid ? dt.toISO() : null;
  }

  // Если число — это может быть Excel serial (включая дробную часть = время)
  if (typeof dateRaw === 'number') {
    // 1899-12-30 + N дней (фракции = часы/мин/сек)
    const base = DateTime.fromISO('1899-12-30', { zone });
    let dt = base.plus({ days: dateRaw });
    // если есть отдельная колонка времени — перезапишем время
    if (timeRaw !== undefined && timeRaw !== null && String(timeRaw).trim() !== '') {
      const t = normalizeTimeFromAny(timeRaw);
      if (t) {
        dt = DateTime.fromISO(`${dt.toFormat('yyyy-LL-dd')}T${t}`, { zone });
      }
    }
    return dt.isValid ? dt.toISO() : null;
  }

  // Если строка — сначала пробуем форматы 'дата время' прямо в одной колонке
  if (typeof dateRaw === 'string') {
    const s = dateRaw.trim();

    // Популярные форматы: "12.09.2025 15:15[:ss]" и ISO-подобные
    const dtCandidates = [
      'dd.LL.yyyy HH:mm:ss',
      'dd.LL.yyyy HH:mm',
      'd.L.yyyy HH:mm:ss',
      'd.L.yyyy HH:mm',
      'yyyy-LL-dd HH:mm:ss',
      'yyyy-LL-dd HH:mm',
    ];
    for (const fmt of dtCandidates) {
      const dt = DateTime.fromFormat(s, fmt, { zone });
      if (dt.isValid) return dt.toISO();
    }

    // Если времени в строке нет — берём дату и пытаемся добавить timeRaw или синтез
    const dateOnlyCandidates = [
      'dd.LL.yyyy',
      'd.L.yyyy',
      'yyyy-LL-dd',
    ];
    for (const fmt of dateOnlyCandidates) {
      const dOnly = DateTime.fromFormat(s, fmt, { zone });
      if (dOnly.isValid) {
        // отдельная колонка времени?
        let timePart = null;
        if (timeRaw !== undefined && timeRaw !== null && String(timeRaw).trim() !== '') {
          timePart = normalizeTimeFromAny(timeRaw);
        }
        if (!timePart) {
          const hour = 9 + (index % 9); // синтез стабильного времени 09..17
          timePart = `${String(hour).padStart(2,'0')}:00:00`;
        }
        const composed = DateTime.fromISO(
          `${dOnly.toFormat('yyyy-LL-dd')}T${timePart}`,
          { zone }
        );
        return composed.isValid ? composed.toISO() : null;
      }
    }

    // На крайний случай — попробуем нативный парсер
    const dt = DateTime.fromJSDate(new Date(s), { zone });
    if (dt.isValid) return dt.toISO();
  }

  // Если ничего не распознали — вернём null
  return null;
}

// вспомогательная: нормализуем timeRaw в "HH:mm:ss"
function normalizeTimeFromAny(timeRaw) {
  if (typeof timeRaw === 'number') {
    const totalSec = Math.round(timeRaw * 24 * 3600); // доля суток
    const hh = String(Math.floor(totalSec / 3600)).padStart(2, '0');
    const mm = String(Math.floor((totalSec % 3600) / 60)).padStart(2, '0');
    const ss = String(totalSec % 60).padStart(2, '0');
    return `${hh}:${mm}:${ss}`;
  } else {
    const s = String(timeRaw).trim();
    // "15:15" или "15:15:42"
    const m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (m) {
      const hh = String(m[1]).padStart(2, '0');
      const mm = String(m[2]).padStart(2, '0');
      const ss = String(m[3] || '00').padStart(2, '0');
      return `${hh}:${mm}:${ss}`;
    }
  }
  return null;
}


// --- конвертер XLS/XLSX -> XLSX (утилита) ---
app.post('/convert/xlsx', upload.single('file'), (req, res) => {
  const rid = req.id;
  const started = Date.now();
  try {
    if (!req.file) return res.status(400).json({ error: 'NO_FILE', requestId: rid });

    const name = (req.file.originalname || '').toLowerCase();
    if (!/\.(xls|xlsx)$/.test(name)) {
      return res.status(400).json({ error: 'ONLY_XLS_OR_XLSX', requestId: rid });
    }

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
    const out = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    const base = (req.file.originalname || 'converted').replace(/\.[^.]+$/, '');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${base}.xlsx"`);

    const parseMs = Date.now() - started;
    console.log(`[convert] rid=${rid} file="${req.file.originalname}" size=${req.file.size}B parseMs=${parseMs}`);
    return res.send(out);
  } catch (e) {
    console.error(`[convert][error] rid=${rid}`, e);
    return res.status(500).json({ error: 'CONVERT_FAILED', requestId: rid });
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
  }
});

app.post('/api/statement/parse', uploadXlsOnly.single('file'), async (req, res) => {
  const rid = req.id;
  const started = Date.now();
  const zone = 'Asia/Bishkek';

  try {
    if (!req.file) return res.status(400).json({ error: 'NO_FILE' });

    const filepath = req.file.path;
    const originalName = req.file.originalname;
    const size = req.file.size;

    const buf = await fs.readFile(filepath);
    const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];

    const isDebug = String(req.query?.debug || '') === '1';
    const HEADER_INDEX = 12; // 13-я строка — как в debug

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
          previewFixed: rowsFixed.slice(0, 3)
        }
      });
    }

    // ---- парсинг транзакций ----
      const transactions = [];
      for (const [i, r] of rowsFixed.entries()) {
        const dateRaw = pick(r, COLS.date);
        const timeRaw = pick(r, COLS.time);
        const tsISO = buildTsISO(dateRaw, timeRaw, i, zone);
        if (!tsISO) continue;


        // MBank: Debit => приход, Credit => расход
        const incomeRaw = pick(r, COLS.debitIsIncomeKeys);    // приход (плюс)
        const expenseRaw = pick(r, COLS.creditIsExpenseKeys); // расход (плюс)

        const income = Number(parseKgsNumber(incomeRaw)) || 0;
        const expense = Number(parseKgsNumber(expenseRaw)) || 0;

        if (income <= 0 && expense <= 0) continue;

        const amount = income - expense; // >0 приход, <0 расход
        const tsDate = DateTime.fromISO(tsISO, { zone }).toJSDate();
          
          const rawDesc = (pick(r, COLS.desc) ?? '').toString();
          const desc = rawDesc.replace(/\\\\/g, '\\').trim();


        transactions.push({
          ts: tsISO,                       // <-- ПОЛНЫЙ ISO С ВРЕМЕНЕМ
          date: ymd(tsDate),               // оставим для совместимости
          description: desc,
          amount,
          credit: income,
          debit: expense,
          direction: amount < 0 ? 'debit' : 'credit'
        });
      }

    // период
    let from = null, to = null;
    if (transactions.length) {
      const dates = transactions.map(t => t.date).sort();
      from = dates[0];
      to = dates[dates.length - 1];
    }

    // агрегаты по дням
    const byDay = new Map();
    for (const t of transactions) {
      const d = byDay.get(t.date) || { date: t.date, credit: 0, debit: 0 };
      if (t.amount > 0) d.credit += t.amount;           // приход
      if (t.amount < 0) d.debit += Math.abs(t.amount);  // расход (по модулю)
      byDay.set(t.date, d);
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
          debit: Number(v.debit.toFixed(2)),
          net: Number((v.credit - v.debit).toFixed(2)),
          // совместимость с текущим фронтом: amount = расходы за день
          amount: Number(v.debit.toFixed(2))
        });
        cur = cur.plus({ days: 1 });
      }
    }

    // totals
    const totals = (() => {
      const credits = transactions.reduce((s, t) => s + (t.credit || 0), 0);
      const debits = transactions.reduce((s, t) => s + (t.debit || 0), 0);
      const net = credits - debits;
      const expenses = Number(debits.toFixed(2));
      return {
        credits: Number(credits.toFixed(2)),
        debits: Number(debits.toFixed(2)),
        net: Number(net.toFixed(2)),
        expenses,                 // новое явное поле
        spending: expenses        // обратная совместимость со старым именем
      };
    })();
      
      // --- cumulative timeline (по каждой транзакции) и end-of-day cumulativeClose ---
      let timeline = [];
      let dailyWithCum = dailySpending;
      if ((transactions?.length || 0) > 0 && from && to) {
        timeline = buildCumulativeTimeline(
          transactions.map(t => ({ ts: t.ts, amount: t.amount })),  // минимально нужные поля
          { tz: zone }
        );

        const dailyCum = attachDailyCumulativeClose(
          { from, to },
          transactions.map(t => ({ ts: t.ts, amount: t.amount })),
          { tz: zone }
        );

        // примиксуем cumulativeClose в dailySpending по дате
        const byDateCum = new Map(dailyCum.map(x => [x.date, x.cumulativeClose]));
        dailyWithCum = dailySpending.map(d => ({
          ...d,
          cumulativeClose: Number((byDateCum.get(d.date) ?? 0).toFixed(2))
        }));
      }


    const parseMs = Date.now() - started;

    // account.*
    const account = { currency: 'KGS', bank: 'MBank' };

    // meta.*
    const meta = {
      processedAt: DateTime.now().setZone(zone).toISO(),
      requestId: rid,
      file: { name: originalName, size },
      sheet: sheetName,
      rows: rowsFixed.length,
      parseMs
    };
      
      meta.contract = {
        timeline: { ts: 'ISO+06:00', cumulative: 'number' },
        dailySpending: { amount: 'expensesPositive', cumulativeClose: 'eod cumulative', net: 'deprecated' }
      };

    // логи (консоль)
    console.log(`[parse] rid=${rid} file="${originalName}" size=${size}B rows=${rowsFixed.length} parseMs=${parseMs}`);

      return res.json({
        meta, account, period: { from, to },
        dailySpending: dailyWithCum,   // <-- с cumulativeClose
        transactions,
        timeline,                      // <-- новая серия для графика ↑/↓
        totals
      });

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
app.use((err, _req, res, next) => {
  if (err?.message === 'ONLY_XLS_ALLOWED') {
    return res.status(415).json({ error: 'ONLY_XLS_ALLOWED' });
  }
  if (err && err instanceof multer.MulterError) {
    if (err.code === 'LIMIT_FILE_SIZE') {
      return res.status(413).json({ error: 'PAYLOAD_TOO_LARGE' });
    }
    return res.status(400).json({ error: 'UPLOAD_ERROR', code: err.code });
  }
  return next(err);
});

// 404
app.use((_req, res) => res.status(404).json({ error: 'Not found' }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`XLS->XLSX converter listening on ${PORT}`));

