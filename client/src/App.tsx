import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { api } from './api';
import { I } from './icons';
import {
  IN_WORK, Order, OrderInput, STATUS_COLORS, STATUS_ORDER, dateInRange, derivedFields, detectCustomer,
  detectMarketplacePlatform, excelDateToIso, extractDateFromName, generateIdentifier,
  normalizeStatus, parseDate, todayStr, workingDays,
} from './utils';

type SortCol =
  | 'zayavka' | 'datePost' | 'mesta' | 'artPlan' | 'artFact' | 'syrPlan' | 'syrFact' | 'status'
  | 'dateProd' | 'dateReady' | 'days' | 'ispolnitel' | 'comment' | 'customer' | 'otchet'
  | 'identifier' | 'dateFromName';

const COLUMNS: { key: SortCol; title: string }[] = [
  { key: 'zayavka', title: 'Заявка' },
  { key: 'datePost', title: 'Дата поступления' },
  { key: 'mesta', title: 'Кол-во мест' },
  { key: 'artPlan', title: 'Артикул (кол-во), план' },
  { key: 'artFact', title: 'Артикул (кол-во), факт' },
  { key: 'syrPlan', title: 'Единиц сырья, план' },
  { key: 'syrFact', title: 'Единиц сырья, факт' },
  { key: 'status', title: 'Статус' },
  { key: 'dateProd', title: 'Передана на производство' },
  { key: 'dateReady', title: 'Дата готовности' },
  { key: 'days', title: 'Время выполнения' },
  { key: 'ispolnitel', title: 'Исполнитель' },
  { key: 'comment', title: 'Комментарий' },
  { key: 'customer', title: 'Заказчик' },
  { key: 'otchet', title: 'Заявки с отчета' },
  { key: 'identifier', title: 'Идентификатор' },
  { key: 'dateFromName', title: 'Дата заявки' },
];

const EXCEL_HEADERS: Record<string, keyof OrderInput> = {
  'заявка': 'zayavka',
  'дата поступления': 'datePost',
  'количество мест': 'mesta',
  'артикул (кол-во) (план)': 'artPlan',
  'артикул (кол-во) (факт)': 'artFact',
  'единиц сырья (план)': 'syrPlan',
  'единиц сырья (факт)': 'syrFact',
  'статус': 'status',
  'дата передачи на производство': 'dateProd',
  'дата готовности': 'dateReady',
  'исполнитель': 'ispolnitel',
  'комментарий': 'comment',
  'заявки с отчета': 'otchet',
  'идентификатор': 'identifier',
};

const DATE_FIELDS = ['datePost', 'dateProd', 'dateReady'];
const PAGE = 100;

interface Filters {
  text: string; status: string; customer: string; ispolnitel: string;
  dateFrom: string; dateTo: string; prodFrom: string; prodTo: string;
  readyFrom: string; readyTo: string; days: string;
}
const EMPTY_FILTERS: Filters = {
  text: '', status: '', customer: '', ispolnitel: '', dateFrom: '', dateTo: '',
  prodFrom: '', prodTo: '', readyFrom: '', readyTo: '', days: '',
};

const EMPTY_FORM = {
  zayavka: '', datePost: '', mesta: '', artPlan: '', artFact: '', syrPlan: '', syrFact: '',
  status: 'Новый', dateProd: '', dateReady: '', ispolnitel: '', comment: '', otchet: '', identifier: '',
};
type Form = typeof EMPTY_FORM;

interface Confirm {
  title: string; message: string; confirmLabel: string; password: boolean;
  run: (password: string) => Promise<void>;
}

function sortValue(o: Order, col: SortCol): string | number {
  switch (col) {
    case 'datePost': return parseDate(o.datePost)?.getTime() ?? 0;
    case 'dateProd': return parseDate(o.dateProd)?.getTime() ?? 0;
    case 'dateReady': return parseDate(o.dateReady)?.getTime() ?? 0;
    case 'mesta': return o.mesta ?? 0;
    case 'syrPlan': return o.syrPlan ?? 0;
    case 'syrFact': return o.syrFact ?? 0;
    case 'status': { const i = STATUS_ORDER.indexOf(o.status); return i === -1 ? 999 : i; }
    case 'days': return workingDays(o.datePost, o.dateReady, o.status) ?? -1;
    case 'customer': return detectCustomer(o.zayavka).toLowerCase();
    case 'dateFromName': return parseDate(extractDateFromName(o.zayavka))?.getTime() ?? 0;
    default: return String((o as any)[col] ?? '').toLowerCase();
  }
}

function plural(n: number, one: string, few: string, many: string) {
  const m10 = n % 10, m100 = n % 100;
  if (m10 === 1 && m100 !== 11) return one;
  if (m10 >= 2 && m10 <= 4 && (m100 < 10 || m100 >= 20)) return few;
  return many;
}

/** Текстовая/числовая ячейка: значение сохраняется при потере фокуса или Enter. */
function TextCell({ value, type = 'text', cls = '', onCommit }: {
  value: string | number | null; type?: string; cls?: string; onCommit: (v: string) => void;
}) {
  const shown = value ?? '';
  return (
    <input
      key={String(shown)}
      type={type}
      className={`cell ${cls}`}
      title={String(shown)}
      defaultValue={shown}
      onBlur={(e) => { if (e.target.value !== String(shown)) onCommit(e.target.value); }}
      onKeyDown={(e) => { if (e.key === 'Enter') (e.target as HTMLInputElement).blur(); }}
    />
  );
}

function DateCell({ value, onCommit }: { value: string | null; onCommit: (v: string) => void }) {
  return <input type="date" className="cell date" value={value ?? ''} onChange={(e) => onCommit(e.target.value)} />;
}

function CustomerGroup({ cls, title, customer, list }: { cls: string; title: string; customer: string; list: Order[] }) {
  const count = (xs: Order[]) => ({
    n: xs.filter((o) => o.status === 'Новый').length,
    w: xs.filter((o) => IN_WORK.includes(o.status)).length,
    d: xs.filter((o) => o.status === 'Выполнен' || o.status === 'Отгружено').length,
    x: xs.filter((o) => o.status === 'Заявка удалена').length,
  });
  let body;
  if (customer === 'маркетплейс') {
    const by: Record<string, Order[]> = {};
    list.forEach((o) => (by[detectMarketplacePlatform(o.zayavka)] ??= []).push(o));
    const entries = Object.entries(by);
    body = entries.length ? entries.map(([p, xs]) => {
      const c = count(xs);
      return (
        <div className="stats-cell" key={p}>
          <span>{p}</span><b>{xs.length}</b>
          <small>нов. {c.n} · в раб. {c.w} · готово {c.d} · удал. {c.x}</small>
        </div>
      );
    }) : <div className="stats-cell"><span>Нет данных</span><b>0</b></div>;
  } else {
    const c = count(list);
    body = ([['Новые', c.n], ['В работе', c.w], ['Выполнено', c.d], ['Удалено', c.x]] as const).map(([name, v]) => (
      <div className="stats-cell" key={name}><span>{name}</span><b>{v}</b></div>
    ));
  }
  return (
    <div className={`stats-group ${cls}`}>
      <h3>{title} <em>{list.length} {plural(list.length, 'заявка', 'заявки', 'заявок')}</em></h3>
      <div className="stats-cells">{body}</div>
    </div>
  );
}

export default function App() {
  const [orders, setOrders] = useState<Order[]>([]);
  const [loading, setLoading] = useState(true);
  const [filters, setFilters] = useState<Filters>(EMPTY_FILTERS);
  const [sort, setSort] = useState<{ col: SortCol | null; dir: 'asc' | 'desc' }>({ col: 'datePost', dir: 'desc' });
  const [selected, setSelected] = useState<Set<number>>(new Set());
  const [showFilters, setShowFilters] = useState(false);
  const [showStats, setShowStats] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const [limit, setLimit] = useState(PAGE);
  const [importOpen, setImportOpen] = useState(false);
  const [importMsg, setImportMsg] = useState<{ text: string; error: boolean } | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const [modal, setModal] = useState<{ id: number | null; form: Form } | null>(null);
  const [confirm, setConfirm] = useState<Confirm | null>(null);
  const [confirmPw, setConfirmPw] = useState('');
  const [confirmErr, setConfirmErr] = useState('');
  const [toast, setToast] = useState<{ msg: string; type: string; show: boolean }>({ msg: '', type: 'success', show: false });
  const toastTimer = useRef<number>();
  const fileInput = useRef<HTMLInputElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);

  const showToast = useCallback((msg: string, type = 'success') => {
    setToast({ msg, type, show: true });
    window.clearTimeout(toastTimer.current);
    toastTimer.current = window.setTimeout(() => setToast((t) => ({ ...t, show: false })), 3500);
  }, []);
  const fail = useCallback((e: unknown) => showToast(e instanceof Error ? e.message : 'Ошибка', 'error'), [showToast]);

  const reload = useCallback(async () => {
    try { setOrders(await api.list()); } catch (e) { fail(e); } finally { setLoading(false); }
  }, [fail]);
  useEffect(() => { reload(); }, [reload]);

  // закрытие меню по клику вне и диалогов по Esc
  useEffect(() => {
    const onDoc = (e: MouseEvent) => { if (menuRef.current && !menuRef.current.contains(e.target as Node)) setMenuOpen(false); };
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return;
      setMenuOpen(false); setModal(null); setConfirm(null); setImportOpen(false);
    };
    document.addEventListener('mousedown', onDoc);
    document.addEventListener('keydown', onKey);
    return () => { document.removeEventListener('mousedown', onDoc); document.removeEventListener('keydown', onKey); };
  }, []);

  useEffect(() => { setLimit(PAGE); }, [filters, sort]);

  // ---------- данные ----------
  const replace = (o: Order) => setOrders((prev) => prev.map((x) => (x.id === o.id ? o : x)));

  async function patch(o: Order, changes: OrderInput) {
    replace({ ...o, ...changes });
    try { replace(await api.update(o.id, changes)); } catch (e) { replace(o); fail(e); }
  }

  function updateField(o: Order, field: keyof OrderInput, value: string) {
    const changes: OrderInput = { [field]: value } as OrderInput;
    if (field === 'zayavka') {
      const id = generateIdentifier(value);
      if (id) changes.identifier = id;
      Object.assign(changes, derivedFields(value));
    }
    if (field === 'mesta' || field === 'syrPlan' || field === 'syrFact') (changes as any)[field] = value === '' ? null : Number(value);
    if (DATE_FIELDS.includes(field)) (changes as any)[field] = value || null;
    patch(o, changes);
  }

  async function updateStatus(o: Order, status: string) {
    await patch(o, { status });
    showToast(`Статус изменён: ${status}`);
  }

  function askConfirm(c: Confirm) { setConfirmPw(''); setConfirmErr(''); setConfirm(c); }

  async function runConfirm() {
    if (!confirm) return;
    try { await confirm.run(confirmPw); setConfirm(null); }
    catch (e) { setConfirmErr(e instanceof Error ? e.message : 'Ошибка'); }
  }

  function removeOrder(o: Order) {
    askConfirm({
      title: 'Удалить заявку?', message: `«${o.zayavka || 'Без названия'}» будет удалена без возможности восстановления.`,
      confirmLabel: 'Удалить', password: false,
      run: async () => {
        await api.remove(o.id);
        setOrders((prev) => prev.filter((x) => x.id !== o.id));
        setSelected((s) => { const n = new Set(s); n.delete(o.id); return n; });
        showToast('Заявка удалена');
      },
    });
  }

  function deleteSelected() {
    const ids = [...selected];
    askConfirm({
      title: `Удалить выбранные заявки (${ids.length})?`, message: 'Действие нельзя отменить. Для подтверждения введите пароль.',
      confirmLabel: 'Удалить', password: true,
      run: async (pw) => {
        const r = await api.deleteMany(ids, pw);
        const set = new Set(ids);
        setOrders((prev) => prev.filter((x) => !set.has(x.id)));
        setSelected(new Set());
        showToast(`Удалено заявок: ${r.deleted}`);
      },
    });
  }

  function clearAll() {
    setMenuOpen(false);
    askConfirm({
      title: 'Очистить весь реестр?', message: `Будут удалены все заявки (${orders.length}). Действие нельзя отменить. Для подтверждения введите пароль.`,
      confirmLabel: 'Удалить всё', password: true,
      run: async (pw) => { await api.clear(pw); setOrders([]); setSelected(new Set()); showToast('Реестр очищен'); },
    });
  }

  // ---------- фильтрация / сортировка ----------
  const filtered = useMemo(() => {
    const f = filters;
    const text = f.text.toLowerCase().trim();
    const list = orders.filter((o) => {
      if (text && !Object.values(o).join(' ').toLowerCase().includes(text)) return false;
      if (f.status && o.status !== f.status) return false;
      if (f.customer && detectCustomer(o.zayavka) !== f.customer) return false;
      if (f.ispolnitel && o.ispolnitel !== f.ispolnitel) return false;
      if ((f.dateFrom || f.dateTo) && !dateInRange(o.datePost, f.dateFrom, f.dateTo)) return false;
      if ((f.prodFrom || f.prodTo) && !dateInRange(o.dateProd, f.prodFrom, f.prodTo)) return false;
      if ((f.readyFrom || f.readyTo) && !dateInRange(o.dateReady, f.readyFrom, f.readyTo)) return false;
      if (f.days) {
        const d = workingDays(o.datePost, o.dateReady, o.status);
        if (d === null) return false;
        if (f.days === 'fast' && d > 5) return false;
        if (f.days === 'slow' && d <= 5) return false;
      }
      return true;
    });
    if (sort.col) {
      const col = sort.col, k = sort.dir === 'asc' ? 1 : -1;
      list.sort((a, b) => { const x = sortValue(a, col), y = sortValue(b, col); return x < y ? -k : x > y ? k : 0; });
    }
    return list;
  }, [orders, filters, sort]);

  const ispolniteli = useMemo(() => [...new Set(orders.map((o) => o.ispolnitel).filter(Boolean))].sort(), [orders]);

  const statusCounts = useMemo(() => {
    const m: Record<string, number> = {};
    orders.forEach((o) => { m[o.status] = (m[o.status] || 0) + 1; });
    return m;
  }, [orders]);

  const advancedActive = useMemo(() => {
    const { text, status, ...rest } = filters;
    return Object.values(rest).filter(Boolean).length;
  }, [filters]);
  const anyFilter = advancedActive > 0 || !!filters.text || !!filters.status;

  const setFilter = (k: keyof Filters) => (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) =>
    setFilters((p) => ({ ...p, [k]: e.target.value }));

  function sortBy(col: SortCol) {
    setSort((s) => (s.col === col ? { col, dir: s.dir === 'asc' ? 'desc' : 'asc' } : { col, dir: 'asc' }));
  }

  function resetFilters() { setFilters(EMPTY_FILTERS); setSort({ col: 'datePost', dir: 'desc' }); }

  // ---------- статистика ----------
  const stats = useMemo(() => {
    const completed = filtered.filter((o) => o.status === 'Выполнен' && o.datePost && o.dateReady);
    const avg = completed.length
      ? (completed.reduce((s, o) => s + (workingDays(o.datePost, o.dateReady, o.status) || 0), 0) / completed.length).toFixed(1)
      : '—';
    return {
      total: filtered.length,
      new: filtered.filter((o) => o.status === 'Новый').length,
      work: filtered.filter((o) => IN_WORK.includes(o.status)).length,
      done: filtered.filter((o) => o.status === 'Выполнен' || o.status === 'Отгружено').length,
      avg,
    };
  }, [filtered]);

  // ---------- форма ----------
  function openModal(o?: Order) {
    if (!o) return setModal({ id: null, form: { ...EMPTY_FORM, datePost: todayStr() } });
    setModal({
      id: o.id,
      form: {
        zayavka: o.zayavka, datePost: o.datePost ?? '', mesta: o.mesta?.toString() ?? '',
        artPlan: o.artPlan, artFact: o.artFact, syrPlan: o.syrPlan?.toString() ?? '',
        syrFact: o.syrFact?.toString() ?? '', status: o.status, dateProd: o.dateProd ?? '',
        dateReady: o.dateReady ?? '', ispolnitel: o.ispolnitel, comment: o.comment,
        otchet: o.otchet, identifier: o.identifier,
      },
    });
  }

  async function saveModal() {
    if (!modal) return;
    const f = modal.form;
    if (!f.zayavka.trim()) return showToast('Введите название заявки', 'error');
    const num = (v: string) => (v === '' ? null : Number(v));
    const data: OrderInput = {
      zayavka: f.zayavka, datePost: f.datePost || null, mesta: num(f.mesta),
      artPlan: f.artPlan, artFact: f.artFact, syrPlan: num(f.syrPlan), syrFact: num(f.syrFact),
      status: f.status, dateProd: f.dateProd || null, dateReady: f.dateReady || null,
      ispolnitel: f.ispolnitel, comment: f.comment, otchet: f.otchet,
      identifier: f.identifier || generateIdentifier(f.zayavka),
      ...derivedFields(f.zayavka),
    };
    try {
      if (modal.id !== null) { replace(await api.update(modal.id, data)); showToast('Заявка сохранена'); }
      else { const created = await api.create(data); setOrders((p) => [...p, created]); showToast('Заявка добавлена'); }
      setModal(null);
    } catch (e) { fail(e); }
  }

  const setForm = (k: keyof Form) => (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) =>
    setModal((m) => (m ? { ...m, form: { ...m.form, [k]: e.target.value } } : m));

  // ---------- Excel ----------
  function templateExcel() {
    setMenuOpen(false);
    const ws = XLSX.utils.aoa_to_sheet([[
      'Заявка', 'Дата поступления', 'Количество мест', 'Артикул (кол-во) (План)', 'Артикул (кол-во) (Факт)',
      'Единиц сырья (План)', 'Единиц сырья (Факт)', 'Статус', 'Дата передачи на производство',
      'Дата готовности', 'Исполнитель', 'Комментарий', 'Заявки с отчета', 'Идентификатор',
    ]]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Заявки');
    XLSX.writeFile(wb, 'Шаблон_НПП_Заявок.xlsx');
  }

  function exportExcel() {
    setMenuOpen(false);
    const rows = filtered.map((o) => ({
      'Заявка': o.zayavka, 'Дата поступления': o.datePost ?? '', 'Количество мест': o.mesta ?? '',
      'Артикул (кол-во) (План)': o.artPlan, 'Артикул (кол-во) (Факт)': o.artFact,
      'Единиц сырья (План)': o.syrPlan ?? '', 'Единиц сырья (Факт)': o.syrFact ?? '', 'Статус': o.status,
      'Дата передачи на производство': o.dateProd ?? '', 'Дата готовности': o.dateReady ?? '',
      'Время выполнения (дни)': workingDays(o.datePost, o.dateReady, o.status) ?? '',
      'Исполнитель': o.ispolnitel, 'Комментарий': o.comment, 'Заказчик': detectCustomer(o.zayavka),
      'Заявки с отчета': o.otchet, 'Идентификатор': o.identifier || generateIdentifier(o.zayavka),
      'Дата заявки': extractDateFromName(o.zayavka),
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Заявки');
    let suffix = '';
    if (filters.status) suffix += '_' + filters.status;
    if (filters.customer) suffix += '_' + filters.customer;
    if (filters.text) suffix += '_поиск';
    XLSX.writeFile(wb, `НПП_Заявки_${todayStr()}${suffix}.xlsx`);
    showToast(`Выгружено заявок: ${rows.length}`);
  }

  async function importFile(file: File) {
    setImportMsg({ text: 'Читаю файл…', error: false });
    try {
      const wb = XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: 'array' });
      const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
      if (rows.length < 2) return setImportMsg({ text: 'Файл пустой или в нём нет строк с данными.', error: true });

      const colMap: Record<number, keyof OrderInput> = {};
      (rows[0] as unknown[]).forEach((h, idx) => {
        const clean = String(h).replace(/[()]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
        for (const [key, field] of Object.entries(EXCEL_HEADERS)) {
          const k = key.replace(/[()]/g, '').replace(/\s+/g, ' ').trim();
          if (clean === k || (clean && (clean.includes(k) || k.includes(clean)))) { colMap[idx] = field; break; }
        }
      });

      const payload: OrderInput[] = [];
      for (const row of rows.slice(1) as unknown[][]) {
        if (!row || !row[0]) continue;
        const o: Record<string, unknown> = { status: 'Новый' };
        row.forEach((cell, idx) => {
          const field = colMap[idx];
          if (!field || cell === '' || cell === undefined) return;
          if (DATE_FIELDS.includes(field)) o[field] = excelDateToIso(cell);
          else if (field === 'status') o[field] = normalizeStatus(String(cell));
          else if (field === 'mesta' || field === 'syrPlan' || field === 'syrFact') o[field] = Number(cell);
          else o[field] = String(cell).trim();
        });
        if (!o.zayavka) continue;
        o.identifier = o.identifier || generateIdentifier(String(o.zayavka));
        Object.assign(o, derivedFields(String(o.zayavka)));
        payload.push(o as OrderInput);
      }

      const r = await api.bulk(payload);
      const skipped = rows.length - 1 - r.imported;
      await reload();
      setImportMsg({ text: `Добавлено заявок: ${r.imported}. Пропущено строк: ${skipped}.`, error: false });
      showToast(`Импортировано заявок: ${r.imported}`);
    } catch (e) {
      setImportMsg({ text: 'Не удалось импортировать: ' + (e instanceof Error ? e.message : e), error: true });
    }
  }

  // ---------- выбор строк ----------
  const allSelected = filtered.length > 0 && filtered.every((o) => selected.has(o.id));
  const toggleAll = () => setSelected(allSelected ? new Set() : new Set(filtered.map((o) => o.id)));
  const toggleOne = (id: number) => setSelected((s) => { const n = new Set(s); n.has(id) ? n.delete(id) : n.add(id); return n; });

  const visible = filtered.slice(0, limit);
  const byCustomer = (c: string) => filtered.filter((o) => detectCustomer(o.zayavka) === c);

  return (
    <>
      <header className="appbar">
        <div className="appbar-inner">
          <div className="brand">
            <h1>Регистр заявок НПП</h1>
            <span>Направление предпродажной подготовки</span>
          </div>
          <div className="search">
            {I.search}
            <input
              type="search" placeholder="Поиск по заявке, артикулу, исполнителю, комментарию"
              aria-label="Поиск" value={filters.text} onChange={setFilter('text')}
            />
          </div>
          <div className="appbar-actions">
            <button className="btn btn-primary" onClick={() => openModal()}>{I.plus}<span className="txt">Добавить заявку</span></button>
            <div className="menu-wrap" ref={menuRef}>
              <button className="btn btn-dark" aria-haspopup="menu" aria-expanded={menuOpen} onClick={() => setMenuOpen((v) => !v)}>
                {I.more}<span className="txt">Ещё</span>
              </button>
              {menuOpen && (
                <div className="menu" role="menu">
                  <button role="menuitem" onClick={() => { setMenuOpen(false); setImportMsg(null); setImportOpen(true); }}>{I.upload}Импорт из Excel</button>
                  <button role="menuitem" onClick={exportExcel}>{I.download}Выгрузить в Excel ({filtered.length})</button>
                  <button role="menuitem" onClick={templateExcel}>{I.file}Скачать шаблон Excel</button>
                  <hr />
                  <button role="menuitem" className="danger" onClick={clearAll}>{I.trash}Очистить весь реестр…</button>
                </div>
              )}
            </div>
          </div>
        </div>
      </header>

      <main className="page">
        <section className="summary" aria-label="Сводка">
          <div><b>{stats.total}</b><span>{plural(stats.total, 'заявка', 'заявки', 'заявок')}{anyFilter ? ' по фильтру' : ''}</span></div>
          <div><b>{stats.new}</b><span>новых</span></div>
          <div><b>{stats.work}</b><span>в работе</span></div>
          <div><b>{stats.done}</b><span>выполнено</span></div>
          <div><b>{stats.avg}</b><span>раб. дней в среднем</span></div>
          <button className="link" onClick={() => setShowStats((v) => !v)}>{showStats ? 'Скрыть разбивку по заказчикам' : 'Разбивка по заказчикам'}</button>
        </section>

        {showStats && (
          <div className="stats-panel">
            <CustomerGroup cls="dlt" title="ДЛТ" customer="ДЛТ" list={byCustomer('ДЛТ')} />
            <CustomerGroup cls="retail" title="Ритейл" customer="ритейл" list={byCustomer('ритейл')} />
            <CustomerGroup cls="marketplace" title="Маркетплейс" customer="маркетплейс" list={byCustomer('маркетплейс')} />
            <CustomerGroup cls="brak" title="Брак" customer="Брак" list={byCustomer('Брак')} />
          </div>
        )}

        <div className="chips" role="group" aria-label="Фильтр по статусу">
          <button className="chip" aria-pressed={!filters.status} onClick={() => setFilters((p) => ({ ...p, status: '' }))}>
            Все <em>{orders.length}</em>
          </button>
          {STATUS_ORDER.map((s) => (
            <button key={s} className={`chip ${STATUS_COLORS[s]}`} aria-pressed={filters.status === s}
              onClick={() => setFilters((p) => ({ ...p, status: p.status === s ? '' : s }))}>
              <i />{s}<em>{statusCounts[s] || 0}</em>
            </button>
          ))}
        </div>

        <div className="toolrow">
          <button className="btn btn-sm" aria-expanded={showFilters} onClick={() => setShowFilters((v) => !v)}>
            {I.filter}Фильтры{advancedActive > 0 && <span className="badge">{advancedActive}</span>}
          </button>
          {anyFilter && <button className="btn btn-sm btn-ghost" onClick={resetFilters}>Сбросить всё</button>}
          <span className="count">Показано {Math.min(limit, filtered.length)} из {filtered.length}</span>
        </div>

        {showFilters && (
          <div className="filters">
            <div className="field"><label htmlFor="f-customer">Заказчик</label>
              <select id="f-customer" value={filters.customer} onChange={setFilter('customer')}>
                <option value="">Все заказчики</option>
                {['ритейл', 'ДЛТ', 'маркетплейс', 'Брак'].map((c) => <option key={c} value={c}>{c}</option>)}
              </select></div>
            <div className="field"><label htmlFor="f-isp">Исполнитель</label>
              <select id="f-isp" value={filters.ispolnitel} onChange={setFilter('ispolnitel')}>
                <option value="">Все исполнители</option>
                {ispolniteli.map((i) => <option key={i} value={i}>{i}</option>)}
              </select></div>
            <div className="field"><label htmlFor="f-days">Время выполнения</label>
              <select id="f-days" value={filters.days} onChange={setFilter('days')}>
                <option value="">Любое</option>
                <option value="fast">Быстрые: до 5 рабочих дней</option>
                <option value="slow">Долгие: больше 5 рабочих дней</option>
              </select></div>
            {([
              ['Дата поступления', 'dateFrom', 'dateTo'],
              ['Передана на производство', 'prodFrom', 'prodTo'],
              ['Дата готовности', 'readyFrom', 'readyTo'],
            ] as const).map(([label, from, to]) => (
              <div className="field" key={label}><span className="lbl">{label}</span>
                <div className="field-range">
                  <input type="date" aria-label={`${label}: с`} value={filters[from]} onChange={setFilter(from)} />
                  <input type="date" aria-label={`${label}: по`} value={filters[to]} onChange={setFilter(to)} />
                </div></div>
            ))}
          </div>
        )}

        <div className="table-card">
          {loading ? (
            <div className="loading">Загружаю заявки…</div>
          ) : filtered.length === 0 ? (
            <div className="empty-state">
              {orders.length === 0 ? (
                <>
                  <h3>В реестре пока нет заявок</h3>
                  <p>Добавьте первую заявку или загрузите список из Excel.</p>
                  <button className="btn btn-primary" onClick={() => openModal()}>{I.plus}Добавить заявку</button>
                </>
              ) : (
                <>
                  <h3>Ничего не найдено</h3>
                  <p>Измените поиск или сбросьте фильтры.</p>
                  <button className="btn" onClick={resetFilters}>Сбросить фильтры</button>
                </>
              )}
            </div>
          ) : (
            <>
              <div className="table-scroll">
                <table>
                  <thead>
                    <tr>
                      <th className="sticky-1">
                        <input type="checkbox" className="checkbox" aria-label="Выбрать все" checked={allSelected} onChange={toggleAll} />
                      </th>
                      {COLUMNS.map((c, i) => (
                        <th key={c.key} className={`sortable ${sort.col === c.key ? 'sorted' : ''} ${i === 0 ? 'sticky-2' : ''}`}
                          aria-sort={sort.col === c.key ? (sort.dir === 'asc' ? 'ascending' : 'descending') : 'none'}
                          onClick={() => sortBy(c.key)}>
                          {c.title}
                          <span className="arrow">{sort.col === c.key ? (sort.dir === 'asc' ? '▲' : '▼') : '⇅'}</span>
                        </th>
                      ))}
                      <th>Действия</th>
                    </tr>
                  </thead>
                  <tbody>
                    {visible.map((o) => {
                      const colorClass = STATUS_COLORS[o.status] || 'status-new';
                      const days = workingDays(o.datePost, o.dateReady, o.status);
                      const customer = detectCustomer(o.zayavka);
                      const text = (field: keyof OrderInput, type = 'text', cls = '') => (
                        <td><TextCell value={(o as any)[field]} type={type} cls={cls} onCommit={(v) => updateField(o, field, v)} /></td>
                      );
                      return (
                        <tr key={o.id} className={`${colorClass} ${selected.has(o.id) ? 'selected' : ''}`}>
                          <td className="sticky-1">
                            <input type="checkbox" className="checkbox" aria-label="Выбрать заявку" checked={selected.has(o.id)} onChange={() => toggleOne(o.id)} />
                          </td>
                          <td className="sticky-2"><TextCell value={o.zayavka} cls="wide" onCommit={(v) => updateField(o, 'zayavka', v)} /></td>
                          <td><DateCell value={o.datePost} onCommit={(v) => updateField(o, 'datePost', v)} /></td>
                          {text('mesta', 'number', 'num')}
                          {text('artPlan', 'text', 'art')}
                          {text('artFact', 'text', 'art')}
                          {text('syrPlan', 'number', 'num')}
                          {text('syrFact', 'number', 'num')}
                          <td>
                            <select className="pill" aria-label="Статус" value={o.status} onChange={(e) => updateStatus(o, e.target.value)}>
                              {STATUS_ORDER.map((s) => <option key={s} value={s}>{s}</option>)}
                            </select>
                          </td>
                          <td><DateCell value={o.dateProd} onCommit={(v) => updateField(o, 'dateProd', v)} /></td>
                          <td><DateCell value={o.dateReady} onCommit={(v) => updateField(o, 'dateReady', v)} /></td>
                          <td>
                            {days !== null
                              ? <span className={`tag ${days <= 5 ? 'good' : 'bad'}`}>{days} {plural(days, 'день', 'дня', 'дней')}</span>
                              : <span className="tag empty" />}
                          </td>
                          {text('ispolnitel')}
                          {text('comment')}
                          <td>{customer ? <span className="tag">{customer}</span> : null}</td>
                          {text('otchet')}
                          <td><TextCell value={o.identifier || generateIdentifier(o.zayavka)} onCommit={(v) => updateField(o, 'identifier', v)} /></td>
                          <td>{extractDateFromName(o.zayavka) && <span className="tag">{extractDateFromName(o.zayavka)}</span>}</td>
                          <td>
                            <div className="row-actions">
                              <button className="icon-btn" title="Редактировать" aria-label="Редактировать" onClick={() => openModal(o)}>{I.edit}</button>
                              <button className="icon-btn danger" title="Удалить" aria-label="Удалить" onClick={() => removeOrder(o)}>{I.trash}</button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              {filtered.length > limit && (
                <div className="more">
                  <button className="btn" onClick={() => setLimit((l) => l + PAGE)}>
                    Показать ещё {Math.min(PAGE, filtered.length - limit)}
                  </button>
                </div>
              )}
            </>
          )}
        </div>
      </main>

      {selected.size > 0 && (
        <div className="selbar" role="region" aria-label="Выбранные заявки">
          <span>Выбрано: <b>{selected.size}</b></span>
          <button className="btn btn-sm btn-dark" onClick={() => setSelected(new Set())}>Снять выбор</button>
          <button className="btn btn-sm btn-danger" onClick={deleteSelected}>{I.trash}Удалить</button>
        </div>
      )}

      {modal && (
        <div className="overlay" onMouseDown={(e) => { if (e.target === e.currentTarget) setModal(null); }}>
          <div className="modal" role="dialog" aria-modal="true" aria-labelledby="m-title">
            <div className="modal-head">
              <div>
                <h2 id="m-title">{modal.id === null ? 'Новая заявка' : 'Редактирование заявки'}</h2>
                <p>Заказчик и дата заявки определяются по названию автоматически.</p>
              </div>
              <button className="icon-btn" aria-label="Закрыть" onClick={() => setModal(null)}>{I.close}</button>
            </div>
            <div className="modal-body">
              <div className="field full"><label htmlFor="m-zayavka">Название заявки</label>
                <input id="m-zayavka" autoFocus type="text" value={modal.form.zayavka} onChange={setForm('zayavka')}
                  placeholder="Например: ДЛТ-Заявка от 15.06.2024" /></div>
              <div className="field"><label htmlFor="m-datePost">Дата поступления</label>
                <input id="m-datePost" type="date" value={modal.form.datePost} onChange={setForm('datePost')} /></div>
              <div className="field"><label htmlFor="m-status">Статус</label>
                <select id="m-status" value={modal.form.status} onChange={setForm('status')}>
                  {STATUS_ORDER.map((s) => <option key={s} value={s}>{s}</option>)}
                </select></div>
              <div className="section">Объём</div>
              <div className="field"><label htmlFor="m-mesta">Количество мест</label>
                <input id="m-mesta" type="number" min={0} value={modal.form.mesta} onChange={setForm('mesta')} /></div>
              <div className="field"><label htmlFor="m-isp">Исполнитель</label>
                <input id="m-isp" type="text" value={modal.form.ispolnitel} onChange={setForm('ispolnitel')} /></div>
              <div className="field"><label htmlFor="m-ap">Артикул (кол-во), план</label>
                <input id="m-ap" type="text" value={modal.form.artPlan} onChange={setForm('artPlan')} /></div>
              <div className="field"><label htmlFor="m-af">Артикул (кол-во), факт</label>
                <input id="m-af" type="text" value={modal.form.artFact} onChange={setForm('artFact')} /></div>
              <div className="field"><label htmlFor="m-sp">Единиц сырья, план</label>
                <input id="m-sp" type="number" min={0} value={modal.form.syrPlan} onChange={setForm('syrPlan')} /></div>
              <div className="field"><label htmlFor="m-sf">Единиц сырья, факт</label>
                <input id="m-sf" type="number" min={0} value={modal.form.syrFact} onChange={setForm('syrFact')} /></div>
              <div className="section">Сроки</div>
              <div className="field"><label htmlFor="m-dp">Передана на производство</label>
                <input id="m-dp" type="date" value={modal.form.dateProd} onChange={setForm('dateProd')} /></div>
              <div className="field"><label htmlFor="m-dr">Дата готовности</label>
                <input id="m-dr" type="date" value={modal.form.dateReady} onChange={setForm('dateReady')} /></div>
              <div className="section">Дополнительно</div>
              <div className="field"><label htmlFor="m-ot">Заявки с отчета</label>
                <input id="m-ot" type="text" value={modal.form.otchet} onChange={setForm('otchet')} /></div>
              <div className="field"><label htmlFor="m-id">Идентификатор</label>
                <input id="m-id" type="text" value={modal.form.identifier} onChange={setForm('identifier')} /></div>
              <div className="field full"><label htmlFor="m-c">Комментарий</label>
                <textarea id="m-c" value={modal.form.comment} onChange={setForm('comment')} /></div>
            </div>
            <div className="modal-foot">
              <button className="btn" onClick={() => setModal(null)}>Отмена</button>
              <button className="btn btn-primary" onClick={saveModal}>Сохранить</button>
            </div>
          </div>
        </div>
      )}

      {confirm && (
        <div className="overlay" onMouseDown={(e) => { if (e.target === e.currentTarget) setConfirm(null); }}>
          <form className="modal small" role="alertdialog" aria-modal="true" aria-labelledby="c-title"
            onSubmit={(e) => { e.preventDefault(); runConfirm(); }}>
            <div className="modal-head"><div><h2 id="c-title">{confirm.title}</h2><p>{confirm.message}</p></div></div>
            {confirm.password && (
              <div className="modal-body single">
                <div className="field"><label htmlFor="c-pw">Пароль</label>
                  <input id="c-pw" type="password" autoFocus autoComplete="off" value={confirmPw}
                    onChange={(e) => { setConfirmPw(e.target.value); setConfirmErr(''); }} /></div>
                {confirmErr && <p className="error-text" role="alert">{confirmErr}</p>}
              </div>
            )}
            {!confirm.password && confirmErr && <div className="modal-body single"><p className="error-text" role="alert">{confirmErr}</p></div>}
            <div className="modal-foot">
              <button type="button" className="btn" onClick={() => setConfirm(null)}>Отмена</button>
              <button type="submit" className="btn btn-danger" autoFocus={!confirm.password}>{confirm.confirmLabel}</button>
            </div>
          </form>
        </div>
      )}

      {importOpen && (
        <div className="overlay" onMouseDown={(e) => { if (e.target === e.currentTarget) setImportOpen(false); }}>
          <div className="modal small" role="dialog" aria-modal="true" aria-labelledby="i-title">
            <div className="modal-head">
              <div><h2 id="i-title">Импорт из Excel</h2><p>Файлы .xlsx, .xls и .csv. Названия колонок как в шаблоне — шаблон можно скачать в меню «Ещё».</p></div>
              <button className="icon-btn" aria-label="Закрыть" onClick={() => setImportOpen(false)}>{I.close}</button>
            </div>
            <div className="modal-body single">
              <div
                className={`dropzone ${dragOver ? 'over' : ''}`} tabIndex={0} role="button"
                onClick={() => fileInput.current?.click()}
                onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') fileInput.current?.click(); }}
                onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
                onDragLeave={() => setDragOver(false)}
                onDrop={(e) => { e.preventDefault(); setDragOver(false); if (e.dataTransfer.files[0]) importFile(e.dataTransfer.files[0]); }}
              >
                <b>Перетащите файл сюда или нажмите, чтобы выбрать</b>
                <span>Новые строки добавятся к существующим заявкам</span>
              </div>
              <input ref={fileInput} type="file" hidden accept=".xlsx,.xls,.csv"
                onChange={(e) => { if (e.target.files?.[0]) importFile(e.target.files[0]); e.target.value = ''; }} />
              {importMsg && <div className={`import-result ${importMsg.error ? 'error' : ''}`} role="status">{importMsg.text}</div>}
            </div>
            <div className="modal-foot"><button className="btn" onClick={() => setImportOpen(false)}>Закрыть</button></div>
          </div>
        </div>
      )}

      <div className={`toast ${toast.type} ${toast.show ? 'show' : ''}`} role="status" aria-live="polite">{toast.msg}</div>
    </>
  );
}
