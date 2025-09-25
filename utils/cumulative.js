// utils/cumulative.js  (ESM)
import { DateTime } from 'luxon';

function round2(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

/**
 * Точки по КАЖДОЙ транзакции: { ts, cumulative }
 * Пополнение ↑, расход ↓. Стартуем от 0.
 */
export function buildCumulativeTimeline(transactions, { tz = 'Asia/Bishkek' } = {}) {
  const sorted = [...transactions].sort(
    (a, b) =>
      DateTime.fromISO(a.ts, { zone: tz }).toMillis() -
      DateTime.fromISO(b.ts, { zone: tz }).toMillis()
  );
  let running = 0;
  return sorted.map(t => {
    running += t.amount; // расходы <0, пополнения >0
    return { ts: t.ts, cumulative: round2(running) };
  });
}

/**
 * Значение cumulative на КОНЕЦ КАЖДОГО ДНЯ периода: { date, cumulativeClose }
 */
export function attachDailyCumulativeClose(period, transactions, { tz = 'Asia/Bishkek' } = {}) {
  const start = DateTime.fromISO(period.from, { zone: tz }).startOf('day');
  const end   = DateTime.fromISO(period.to,   { zone: tz }).startOf('day');

  const sorted = [...transactions].sort(
    (a, b) =>
      DateTime.fromISO(a.ts, { zone: tz }).toMillis() -
      DateTime.fromISO(b.ts, { zone: tz }).toMillis()
  );

  const result = [];
  let running = 0;
  let ti = 0;

  for (let d = start; d <= end; d = d.plus({ days: 1 })) {
    const dayEnd = d.endOf('day');
    while (ti < sorted.length) {
      const ts = DateTime.fromISO(sorted[ti].ts, { zone: tz });
      if (ts.toMillis() <= dayEnd.toMillis()) {
        running += sorted[ti].amount;
        ti += 1;
      } else {
        break;
      }
    }
    result.push({ date: d.toISODate(), cumulativeClose: round2(running) });
  }
  return result;
}
