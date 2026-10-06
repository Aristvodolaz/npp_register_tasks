export interface Order {
  id: number;
  zayavka: string;
  datePost: string | null;
  mesta: number | null;
  artPlan: string;
  artFact: string;
  syrPlan: number | null;
  syrFact: number | null;
  status: string;
  dateProd: string | null;
  dateReady: string | null;
  ispolnitel: string;
  comment: string;
  otchet: string;
  identifier: string;
  customer: string;
  dateFromName: string;
}

export type OrderInput = Partial<Omit<Order, 'id'>>;

export const STATUS_ORDER = [
  'Новый', 'В работе на приемке', 'Передан на производство', 'В работе на производстве',
  'В работе у операторов', 'Выполнен', 'Отгружено', 'Заявка удалена',
];

export const STATUS_COLORS: Record<string, string> = {
  'Новый': 'status-new',
  'В работе на приемке': 'status-priemka',
  'Передан на производство': 'status-proizvodstvo-passed',
  'В работе на производстве': 'status-proizvodstvo-work',
  'В работе у операторов': 'status-operators',
  'Выполнен': 'status-done',
  'Отгружено': 'status-shipped',
  'Заявка удалена': 'status-deleted',
};

export const IN_WORK = ['В работе на приемке', 'Передан на производство', 'В работе на производстве', 'В работе у операторов'];

const STATUS_NORM: Record<string, string> = {
  'новый': 'Новый', 'новая': 'Новый', 'new': 'Новый',
  'в работе на приемке': 'В работе на приемке', 'приемка': 'В работе на приемке',
  'передан на производство': 'Передан на производство', 'на производство': 'Передан на производство',
  'производство передан': 'Передан на производство',
  'в работе на производстве': 'В работе на производстве', 'производство': 'В работе на производстве',
  'в работе у операторов': 'В работе у операторов', 'операторы': 'В работе у операторов',
  'выполнен': 'Выполнен', 'выполнена': 'Выполнен', 'готово': 'Выполнен', 'done': 'Выполнен',
  'отгружено': 'Отгружено', 'отгружен': 'Отгружено', 'отправлено': 'Отгружено', 'shipped': 'Отгружено',
  'заявка удалена': 'Заявка удалена', 'удалена': 'Заявка удалена', 'удален': 'Заявка удалена', 'deleted': 'Заявка удалена',
};

export function normalizeStatus(raw: string): string {
  const v = raw.toLowerCase().replace(/\s+/g, ' ').trim();
  if (STATUS_NORM[v]) return STATUS_NORM[v];
  for (const [k, norm] of Object.entries(STATUS_NORM)) {
    if (v && (v.includes(k) || k.includes(v))) return norm;
  }
  return 'Новый';
}

/** Парсит ISO (YYYY-MM-DD) и DD.MM.YYYY, возвращает локальную дату. */
export function parseDate(d: string | null | undefined): Date | null {
  if (!d) return null;
  let m = d.match(/^(\d{4})-(\d{2})-(\d{2})/);
  let y: number, mo: number, day: number;
  if (m) { y = +m[1]; mo = +m[2]; day = +m[3]; }
  else if ((m = d.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/))) { day = +m[1]; mo = +m[2]; y = +m[3]; }
  else return null;
  const date = new Date(y, mo - 1, day);
  return date.getDate() === day && date.getMonth() === mo - 1 ? date : null;
}

export function todayStr(): string {
  const d = new Date();
  const p = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

/** Рабочие дни между датами включительно; для удалённых заявок — null. */
export function workingDays(start: string | null, end: string | null, status: string): number | null {
  if (status === 'Заявка удалена') return null;
  const s = parseDate(start);
  const e = parseDate(end);
  if (!s || !e || e < s) return null;
  let days = 0;
  const cur = new Date(s);
  while (cur <= e) {
    const dow = cur.getDay();
    if (dow !== 0 && dow !== 6) days++;
    cur.setDate(cur.getDate() + 1);
  }
  return days;
}

export function detectCustomer(name: string): string {
  const n = (name || '').trim();
  if (!n) return '';
  if (/брак/i.test(n)) return 'Брак';
  if (/^ДЛТ/i.test(n)) return 'ДЛТ';
  if (/Ozon|WB|ЯМ|Яндекс|Wildberries/i.test(n)) return 'маркетплейс';
  if (/^[0-9]/.test(n)) return 'ритейл';
  return '';
}

export function detectMarketplacePlatform(name: string): string {
  const n = (name || '').trim().toLowerCase();
  if (/\bwb\b/.test(n)) return 'WB';
  if (/\bozon\b/.test(n)) return 'Ozon';
  if (/\bям\b|\bяндекс\b|\byandex\b/.test(n)) return 'ЯМ';
  return 'Другое';
}

export function extractDateFromName(name: string): string {
  const n = name || '';
  const p = (s: string) => s.padStart(2, '0');
  let m = n.match(/([0-9]{2})[./-]([0-9]{2})[./-]([0-9]{4})/);
  if (m) return `${m[1]}.${m[2]}.${m[3]}`;
  m = n.match(/([0-9]{4})[./-]([0-9]{2})[./-]([0-9]{2})/);
  if (m) return `${m[3]}.${m[2]}.${m[1]}`;
  m = n.match(/([0-9]{1,2})[./-]([0-9]{1,2})[./-]([0-9]{4})/);
  if (m) return `${p(m[1])}.${p(m[2])}.${m[3]}`;
  const year = new Date().getFullYear();
  m = n.match(/([0-9]{2})[./-]([0-9]{2})(?![./-][0-9])/);
  if (m) return `${m[1]}.${m[2]}.${year}`;
  m = n.match(/([0-9]{1,2})[./-]([0-9]{1,2})(?![./-][0-9])/);
  if (m) return `${p(m[1])}.${p(m[2])}.${year}`;
  return '';
}

export function generateIdentifier(name: string): string {
  const n = (name || '').trim();
  if (/WB/i.test(n)) return 'WB';
  if (/Ozon/i.test(n)) return 'ozon';
  if (/ДЛТ/i.test(n)) return 'ДЛТ';
  const m = n.match(/^(\d+)/);
  return m ? m[1] : '';
}

export function dateInRange(date: string | null, from: string, to: string): boolean {
  const d = parseDate(date);
  if (!d) return false;
  const f = parseDate(from);
  if (f && d < f) return false;
  const t = parseDate(to);
  if (t && d > t) return false;
  return true;
}

/** Значение ячейки Excel -> ISO-дата (серийные номера Excel и DD.MM.YYYY). */
export function excelDateToIso(v: unknown): string | null {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number' && v > 20000 && v < 80000) {
    const dt = new Date(Math.round((v - 25569) * 86400 * 1000));
    return dt.toISOString().slice(0, 10);
  }
  const d = parseDate(String(v).trim());
  if (!d) return null;
  const p = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

export function formatDateRu(iso: string | null): string {
  const d = parseDate(iso);
  return d ? d.toLocaleDateString('ru-RU') : '';
}

/** Производные поля, которые хранятся в БД рядом с названием заявки. */
export function derivedFields(zayavka: string) {
  return { customer: detectCustomer(zayavka), dateFromName: extractDateFromName(zayavka) };
}
