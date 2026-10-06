import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { api } from './api';
import {
  IN_WORK, Order, OrderInput, STATUS_COLORS, STATUS_ORDER, dateInRange, detectCustomer,
  derivedFields, detectMarketplacePlatform, excelDateToIso, extractDateFromName, generateIdentifier,
  normalizeStatus, parseDate, todayStr, workingDays,
} from './utils';

type SortCol =
  | 'zayavka' | 'datePost' | 'mesta' | 'artPlan' | 'artFact' | 'syrPlan' | 'syrFact' | 'status'
  | 'dateProd' | 'dateReady' | 'days' | 'ispolnitel' | 'comment' | 'customer' | 'otchet'
  | 'identifier' | 'dateFromName';

const COLUMNS: { key: SortCol; title: string; style?: React.CSSProperties }[] = [
  { key: 'zayavka', title: 'Заявка', style: { minWidth: 280 } },
  { key: 'datePost', title: 'Дата поступления', style: { width: 110 } },
  { key: 'mesta', title: 'Количество мест', style: { width: 80 } },
  { key: 'artPlan', title: 'Артикул (кол-во) (План)' },
  { key: 'artFact', title: 'Артикул (кол-во) (Факт)' },
  { key: 'syrPlan', title: 'Единиц сырья (План)' },
  { key: 'syrFact', title: 'Единиц сырья (Факт)' },
  { key: 'status', title: 'Статус', style: { width: 180 } },
  { key: 'dateProd', title: 'Дата передачи на производство', style: { width: 140 } },
  { key: 'dateReady', title: 'Дата готовности', style: { width: 110 } },
  { key: 'days', title: 'Время выполнения (дни)', style: { width: 90 } },
  { key: 'ispolnitel', title: 'Исполнитель' },
  { key: 'comment', title: 'Комментарий' },
  { key: 'customer', title: 'Заказчик' },
  { key: 'otchet', title: 'Заявки с отчета' },
  { key: 'identifier', title: 'Идентификатор' },
  { key: 'dateFromName', title: 'Дата заявки', style: { width: 100 } },
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

/** Текстовая/числовая ячейка: сохраняется при потере фокуса. */
function TextCell({ value, type = 'text', cls = '', style, onCommit }: {
  value: string | number | null; type?: string; cls?: string; style?: React.CSSProperties;
  onCommit: (v: string) => void;
}) {
  const shown = value ?? '';
  return (
    <input
      key={String(shown)}
      type={type}
      className={`cell-input ${cls}`}
      style={style}
      defaultValue={shown}
      onBlur={(e) => { if (e.target.value !== String(shown)) onCommit(e.target.value); }}
      onKeyDown={(e) => { if (e.key === 'Enter') (e.target as HTMLInputElement).blur(); }}
    />
  );
}

function DateCell({ value, onCommit }: { value: string | null; onCommit: (v: string) => void }) {
  return (
    <input
      type="date"
      className="cell-input date"
      value={value ?? ''}
      onChange={(e) => onCommit(e.target.value)}
    />
  );
}

export default function App() {
  const [orders, setOrders] = useState<Order[]>([]);
  const [loading, setLoading] = useState(true);
  const [filters, setFilters] = useState<Filters>(EMPTY_FILTERS);
  const [sort, setSort] = useState<{ col: SortCol | null; dir: 'asc' | 'desc' }>({ col: 'datePost', dir: 'desc' });
  const [selected, setSelected] = useState<Set<number>>(new Set());
  const [filtersCollapsed, setFiltersCollapsed] = useState(true);
  const [statsCollapsed, setStatsCollapsed] = useState(true);
  const [importOpen, setImportOpen] = useState(false);
  const [importMsg, setImportMsg] = useState<{ text: string; error: boolean } | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const [modal, setModal] = useState<{ id: number | null; form: Form } | null>(null);
  const [toast, setToast] = useState<{ msg: string; type: string; show: boolean }>({ msg: '', type: 'success', show: false });
  const toastTimer = useRef<number>();
  const fileInput = useRef<HTMLInputElement>(null);
  const tableInner = useRef<HTMLDivElement>(null);

  const showToast = useCallback((msg: string, type = 'success') => {
    setToast({ msg, type, show: true });
    window.clearTimeout(toastTimer.current);
    toastTimer.current = window.setTimeout(() => setToast((t) => ({ ...t, show: false })), 3000);
  }, []);

  const fail = useCallback((e: unknown) => showToast(e instanceof Error ? e.message : 'Ошибка', 'error'), [showToast]);

  const reload = useCallback(async () => {
    try { setOrders(await api.list()); } catch (e) { fail(e); } finally { setLoading(false); }
  }, [fail]);

  useEffect(() => { reload(); }, [reload]);

  // Как в оригинале: прокрутка вниз к последним записям после первой загрузки
  useEffect(() => {
    if (!loading && tableInner.current) tableInner.current.scrollTop = tableInner.current.scrollHeight;
  }, [loading]);

  // ---------- данные ----------
  const replace = (o: Order) => setOrders((prev) => prev.map((x) => (x.id === o.id ? o : x)));

  async function patch(o: Order, changes: OrderInput) {
    const prev = o;
    replace({ ...o, ...changes });
    try {
      const saved = await api.update(o.id, changes);
      replace(saved);
    } catch (e) {
      replace(prev);
      fail(e);
    }
  }

  function updateField(o: Order, field: keyof OrderInput, value: string) {
    const changes: OrderInput = { [field]: value } as OrderInput;
    if (field === 'zayavka') {
      const id = generateIdentifier(value);
      if (id) changes.identifier = id;
      Object.assign(changes, derivedFields(value));
    }
    if (field === 'mesta' || field === 'syrPlan' || field === 'syrFact') {
      (changes as any)[field] = value === '' ? null : Number(value);
    }
    if (DATE_FIELDS.includes(field)) (changes as any)[field] = value || null;
    patch(o, changes);
  }

  async function updateStatus(o: Order, status: string) {
    await patch(o, { status });
    showToast(`Статус изменен на "${status}"`);
  }

  async function removeOrder(o: Order) {
    if (!confirm('Удалить заявку?')) return;
    try {
      await api.remove(o.id);
      setOrders((prev) => prev.filter((x) => x.id !== o.id));
      showToast('Заявка удалена');
    } catch (e) { fail(e); }
  }

  async function deleteSelected() {
    if (selected.size === 0) return showToast('Ни одна заявка не выделена', 'error');
    if (!confirm(`Удалить ${selected.size} выделенных заявок?`)) return;
    const password = prompt('Введите пароль для удаления:');
    if (password === null) return;
    try {
      const ids = [...selected];
      const r = await api.deleteMany(ids, password);
      setOrders((prev) => prev.filter((x) => !selected.has(x.id)));
      setSelected(new Set());
      showToast(`Удалено ${r.deleted} заявок`);
    } catch (e) { fail(e); }
  }

  async function clearAll() {
    const password = prompt('Введите пароль для очистки всех заявок:');
    if (password === null) return;
    if (!confirm('Очистить ВСЕ заявки? Это действие нельзя отменить!')) return;
    try {
      await api.clear(password);
      setOrders([]);
      setSelected(new Set());
      showToast('Все заявки удалены');
    } catch (e) { fail(e); }
  }

  // ---------- фильтрация / сортировка ----------
  const filtered = useMemo(() => {
    const f = filters;
    const text = f.text.toLowerCase();
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
      const col = sort.col;
      const k = sort.dir === 'asc' ? 1 : -1;
      list.sort((a, b) => {
        const x = sortValue(a, col), y = sortValue(b, col);
        return x < y ? -k : x > y ? k : 0;
      });
    }
    return list;
  }, [orders, filters, sort]);

  const ispolniteli = useMemo(
    () => [...new Set(orders.map((o) => o.ispolnitel).filter(Boolean))].sort(),
    [orders],
  );

  const setFilter = (k: keyof Filters) => (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) =>
    setFilters((p) => ({ ...p, [k]: e.target.value }));

  function sortBy(col: SortCol) {
    setSort((s) => (s.col === col ? { col, dir: s.dir === 'asc' ? 'desc' : 'asc' } : { col, dir: 'asc' }));
  }

  function resetFilters() {
    setFilters(EMPTY_FILTERS);
    setSort({ col: null, dir: 'asc' });
    showToast('Фильтры сброшены');
  }

  // ---------- статистика ----------
  const stats = useMemo(() => {
    const completed = filtered.filter((o) => o.status === 'Выполнен' && o.datePost && o.dateReady);
    const avg = completed.length
      ? (completed.reduce((s, o) => s + (workingDays(o.datePost, o.dateReady, o.status) || 0), 0) / completed.length).toFixed(1)
      : '0';
    return {
      total: filtered.length,
      new: filtered.filter((o) => o.status === 'Новый').length,
      work: filtered.filter((o) => IN_WORK.includes(o.status)).length,
      done: filtered.filter((o) => o.status === 'Выполнен' || o.status === 'Отгружено').length,
      avg,
    };
  }, [filtered]);

  // ---------- модалка ----------
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
      if (modal.id !== null) {
        replace(await api.update(modal.id, data));
        showToast('Заявка обновлена');
      } else {
        const created = await api.create(data);
        setOrders((prev) => [...prev, created]);
        showToast('Заявка добавлена');
      }
      setModal(null);
    } catch (e) { fail(e); }
  }

  const setForm = (k: keyof Form) => (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) =>
    setModal((m) => (m ? { ...m, form: { ...m.form, [k]: e.target.value } } : m));

  // ---------- Excel ----------
  function templateExcel() {
    const ws = XLSX.utils.aoa_to_sheet([[
      'Заявка', 'Дата поступления', 'Количество мест', 'Артикул (кол-во) (План)', 'Артикул (кол-во) (Факт)',
      'Единиц сырья (План)', 'Единиц сырья (Факт)', 'Статус', 'Дата передачи на производство',
      'Дата готовности', 'Исполнитель', 'Комментарий', 'Заявки с отчета', 'Идентификатор',
    ]]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Заявки');
    XLSX.writeFile(wb, 'Шаблон_НПП_Заявок.xlsx');
    showToast('Шаблон скачан');
  }

  function exportExcel() {
    const rows = filtered.map((o) => ({
      'Заявка': o.zayavka,
      'Дата поступления': o.datePost ?? '',
      'Количество мест': o.mesta ?? '',
      'Артикул (кол-во) (План)': o.artPlan,
      'Артикул (кол-во) (Факт)': o.artFact,
      'Единиц сырья (План)': o.syrPlan ?? '',
      'Единиц сырья (Факт)': o.syrFact ?? '',
      'Статус': o.status,
      'Дата передачи на производство': o.dateProd ?? '',
      'Дата готовности': o.dateReady ?? '',
      'Время выполнения (дни)': workingDays(o.datePost, o.dateReady, o.status) ?? '',
      'Исполнитель': o.ispolnitel,
      'Комментарий': o.comment,
      'Заказчик': detectCustomer(o.zayavka),
      'Заявки с отчета': o.otchet,
      'Идентификатор': o.identifier || generateIdentifier(o.zayavka),
      'Дата заявки': extractDateFromName(o.zayavka),
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Заявки');
    let suffix = '';
    if (filters.status) suffix += '_' + filters.status;
    if (filters.customer) suffix += '_' + filters.customer;
    if (filters.text) suffix += '_поиск';
    XLSX.writeFile(wb, `НПП_Заявки_${todayStr()}${suffix}.xlsx`);
    showToast(`Экспортировано ${rows.length} заявок`);
  }

  async function importFile(file: File) {
    setImportMsg({ text: 'Чтение файла...', error: false });
    try {
      const wb = XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: 'array' });
      const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
      if (rows.length < 2) return setImportMsg({ text: 'Файл пустой или не содержит данных', error: true });

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
      setImportMsg({ text: `Импорт завершен: ${r.imported} заявок добавлено, ${skipped} пропущено`, error: false });
      showToast(`Импортировано ${r.imported} заявок`);
    } catch (e) {
      setImportMsg({ text: 'Ошибка импорта: ' + (e instanceof Error ? e.message : e), error: true });
      showToast('Ошибка импорта', 'error');
    }
  }

  // ---------- статистика по заказчикам ----------
  function CustomerGroup({ cls, title, customer }: { cls: string; title: string; customer: string }) {
    const list = filtered.filter((o) => detectCustomer(o.zayavka) === customer);
    const count = (xs: Order[]) => ({
      n: xs.filter((o) => o.status === 'Новый').length,
      w: xs.filter((o) => IN_WORK.includes(o.status)).length,
      d: xs.filter((o) => o.status === 'Выполнен' || o.status === 'Отгружено').length,
      x: xs.filter((o) => o.status === 'Заявка удалена').length,
    });
    let body;
    if (customer === 'маркетплейс') {
      const byPlatform: Record<string, Order[]> = {};
      list.forEach((o) => (byPlatform[detectMarketplacePlatform(o.zayavka)] ??= []).push(o));
      const entries = Object.entries(byPlatform);
      body = entries.length ? entries.map(([p, xs]) => {
        const c = count(xs);
        return (
          <div className="subcustomer-card" key={p}>
            <div className="name">{p}</div>
            <div className="value">{xs.length}</div>
            <div className="label">Н:{c.n} | Р:{c.w} | В:{c.d} | У:{c.x}</div>
          </div>
        );
      }) : <div className="subcustomer-card"><div className="name">Нет данных</div><div className="value">0</div></div>;
    } else {
      const c = count(list);
      body = [['Новые', c.n], ['В работе', c.w], ['Выполнено', c.d], ['Удалено', c.x]].map(([name, v]) => (
        <div className="subcustomer-card" key={name}><div className="name">{name}</div><div className="value">{v}</div></div>
      ));
    }
    return (
      <div className={`customer-group ${cls}`}>
        <div className="customer-group-title">{title} <span className="count">{list.length} заявок</span></div>
        <div className="subcustomer-grid">{body}</div>
      </div>
    );
  }

  const allVisibleSelected = filtered.length > 0 && filtered.every((o) => selected.has(o.id));
  const toggleSelectAll = () =>
    setSelected(allVisibleSelected ? new Set() : new Set(filtered.map((o) => o.id)));
  const toggleSelect = (id: number) =>
    setSelected((s) => { const n = new Set(s); n.has(id) ? n.delete(id) : n.add(id); return n; });

  const sortCls = (col: SortCol) => (sort.col === col ? (sort.dir === 'asc' ? 'sort-asc' : 'sort-desc') : '');

  return (
    <div className="container">
      <h1>📋 НПП — Регистр заявок (Направление предпродажной подготовки)</h1>

      <div className="stats">
        <div className="stat-card"><div className="number">{stats.total}</div><div className="label">Всего заявок</div></div>
        <div className="stat-card"><div className="number">{stats.new}</div><div className="label">Новые</div></div>
        <div className="stat-card"><div className="number">{stats.work}</div><div className="label">В работе</div></div>
        <div className="stat-card"><div className="number">{stats.done}</div><div className="label">Выполнено</div></div>
        <div className="stat-card"><div className="number">{stats.avg}</div><div className="label">Среднее время (дн)</div></div>
      </div>

      <div className="customer-stats">
        <div className="customer-stats-header" onClick={() => setStatsCollapsed((v) => !v)}>
          <span className={`arrow ${statsCollapsed ? 'collapsed' : ''}`}>▼</span>
          <span>📊 Статистика по заказчикам</span>
        </div>
        <div className={`customer-stats-content ${statsCollapsed ? 'collapsed' : ''}`}>
          <CustomerGroup cls="dlt" title="🏢 ДЛТ" customer="ДЛТ" />
          <CustomerGroup cls="retail" title="🏪 РИТЕЙЛ" customer="ритейл" />
          <CustomerGroup cls="marketplace" title="🛒 МАРКЕТПЛЕЙС" customer="маркетплейс" />
          <CustomerGroup cls="brak" title="⚠️ БРАК" customer="Брак" />
        </div>
      </div>

      <div className="toolbar" style={{ marginTop: 15 }}>
        <button className="btn-add" onClick={() => openModal()}>➕ Добавить заявку</button>
        <button className="btn-import" onClick={() => setImportOpen((v) => !v)}>📥 Импорт из Excel</button>
        <button className="btn-template" onClick={templateExcel}>📄 Шаблон Excel</button>
        <button className="btn-export" onClick={exportExcel}>📤 Экспорт в Excel</button>
        <button className="btn-delete" style={{ padding: '8px 16px', borderRadius: 6, fontSize: 13 }} onClick={deleteSelected}>🗑️ Удалить выделенные</button>
        <button className="btn-clear" onClick={clearAll}>🗑️ Очистить все</button>
        <button className="btn-reset" onClick={resetFilters}>🔄 Сбросить фильтры</button>
      </div>

      <div className={`import-area ${importOpen ? 'active' : ''}`}>
        <h3 style={{ fontSize: 14, color: '#374151', marginBottom: 10 }}>Импорт заявок из Excel</h3>
        <div
          className={`import-dropzone ${dragOver ? 'dragover' : ''}`}
          onClick={() => fileInput.current?.click()}
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => { e.preventDefault(); setDragOver(false); if (e.dataTransfer.files[0]) importFile(e.dataTransfer.files[0]); }}
        >
          <div className="icon">📁</div>
          <div className="text">Нажмите или перетащите Excel-файл сюда</div>
          <div className="subtext">Поддерживаются форматы .xlsx, .xls, .csv</div>
        </div>
        <input
          ref={fileInput} type="file" className="import-file-input" accept=".xlsx,.xls,.csv"
          onChange={(e) => { if (e.target.files?.[0]) importFile(e.target.files[0]); e.target.value = ''; }}
        />
        {importMsg && <div className={`import-results active ${importMsg.error ? 'error' : ''}`}>{importMsg.text}</div>}
      </div>

      <div className="filters-panel">
        <h3 onClick={() => setFiltersCollapsed((v) => !v)} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <span>🔍 Фильтры и сортировка</span>
          <span id="filters-toggle-icon" className={filtersCollapsed ? 'collapsed' : ''} style={{ fontSize: 12 }}>▼</span>
        </h3>
        <div className={`filters-grid ${filtersCollapsed ? 'collapsed' : ''}`} id="filters-content">
          <div className="filter-group"><label>Поиск по тексту</label>
            <input type="text" placeholder="Введите текст..." value={filters.text} onChange={setFilter('text')} /></div>
          <div className="filter-group"><label>Статус</label>
            <select value={filters.status} onChange={setFilter('status')}>
              <option value="">Все статусы</option>
              {STATUS_ORDER.map((s) => <option key={s} value={s}>{s}</option>)}
            </select></div>
          <div className="filter-group"><label>Заказчик</label>
            <select value={filters.customer} onChange={setFilter('customer')}>
              <option value="">Все заказчики</option>
              {['ритейл', 'ДЛТ', 'маркетплейс', 'Брак'].map((c) => <option key={c} value={c}>{c}</option>)}
            </select></div>
          <div className="filter-group"><label>Исполнитель</label>
            <select value={filters.ispolnitel} onChange={setFilter('ispolnitel')}>
              <option value="">Все исполнители</option>
              {ispolniteli.map((i) => <option key={i} value={i}>{i}</option>)}
            </select></div>
          {([
            ['Дата поступления', 'dateFrom', 'dateTo'],
            ['Дата передачи на производство', 'prodFrom', 'prodTo'],
            ['Дата готовности', 'readyFrom', 'readyTo'],
          ] as const).map(([label, from, to]) => (
            <div className="filter-group" key={label}><label>{label}</label>
              <div className="period-row">
                <span className="date-label">с</span><input type="date" value={filters[from]} onChange={setFilter(from)} />
                <span className="date-label">по</span><input type="date" value={filters[to]} onChange={setFilter(to)} />
              </div></div>
          ))}
          <div className="filter-group"><label>Время выполнения (дни)</label>
            <select value={filters.days} onChange={setFilter('days')}>
              <option value="">Все</option>
              <option value="fast">Быстрые (≤ 5 дней)</option>
              <option value="slow">Долгие (&gt; 5 дней)</option>
            </select></div>
        </div>
      </div>

      <div className="table-outer">
        <div className="table-inner" ref={tableInner}>
          <table id="main-table">
            <thead>
              <tr>
                <th style={{ width: 30 }}>
                  <input type="checkbox" className="checkbox" checked={allVisibleSelected} onChange={toggleSelectAll} />
                </th>
                {COLUMNS.map((c) => (
                  <th key={c.key} style={c.style} className={sortCls(c.key)} onClick={() => sortBy(c.key)}>
                    {c.title}<span className="sort-icon">⇅</span>
                  </th>
                ))}
                <th style={{ width: 80 }}>Действия</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((o) => {
                const colorClass = STATUS_COLORS[o.status] || 'status-new';
                const days = workingDays(o.datePost, o.dateReady, o.status);
                const daysClass = days === null ? '' : days <= 5 ? 'green' : 'red';
                const identifier = o.identifier || generateIdentifier(o.zayavka);
                const text = (field: keyof OrderInput, type = 'text', cls = '') => (
                  <td><TextCell value={(o as any)[field]} type={type} cls={cls} onCommit={(v) => updateField(o, field, v)} /></td>
                );
                return (
                  <tr key={o.id} className={colorClass}>
                    <td><input type="checkbox" className="checkbox" checked={selected.has(o.id)} onChange={() => toggleSelect(o.id)} /></td>
                    <td style={{ minWidth: 280 }}>
                      <TextCell value={o.zayavka} style={{ minWidth: 260 }} onCommit={(v) => updateField(o, 'zayavka', v)} />
                    </td>
                    <td><DateCell value={o.datePost} onCommit={(v) => updateField(o, 'datePost', v)} /></td>
                    {text('mesta', 'number', 'number')}
                    {text('artPlan')}
                    {text('artFact')}
                    {text('syrPlan', 'number', 'number')}
                    {text('syrFact', 'number', 'number')}
                    <td>
                      <select className={`status-select ${colorClass}`} value={o.status} onChange={(e) => updateStatus(o, e.target.value)}>
                        {STATUS_ORDER.map((s) => <option key={s} value={s}>{s}</option>)}
                      </select>
                    </td>
                    <td><DateCell value={o.dateProd} onCommit={(v) => updateField(o, 'dateProd', v)} /></td>
                    <td><DateCell value={o.dateReady} onCommit={(v) => updateField(o, 'dateReady', v)} /></td>
                    <td><span className={`cell-auto ${daysClass}`}>{days !== null ? `${days} дн` : ''}</span></td>
                    {text('ispolnitel')}
                    {text('comment')}
                    <td><span className="cell-auto">{detectCustomer(o.zayavka)}</span></td>
                    {text('otchet')}
                    <td><TextCell value={identifier} onCommit={(v) => updateField(o, 'identifier', v)} /></td>
                    <td><span className="cell-auto">{extractDateFromName(o.zayavka)}</span></td>
                    <td className="actions">
                      <button className="btn-edit" onClick={() => openModal(o)}>✏️</button>
                      <button className="btn-delete" onClick={() => removeOrder(o)}>🗑️</button>
                    </td>
                  </tr>
                );
              })}
              {Array.from({ length: Math.max(0, 10 - filtered.length) }, (_, i) => (
                <tr key={`e${i}`} className="empty-row">
                  <td colSpan={COLUMNS.length + 2} style={{ borderBottom: '1px dashed #e5e7eb' }} />
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {modal && (
        <div className="modal-overlay active" onMouseDown={(e) => { if (e.target === e.currentTarget) setModal(null); }}>
          <div className="modal">
            <h2>{modal.id === null ? 'Добавить заявку' : 'Редактировать заявку'}</h2>
            {([
              ['Название заявки', 'zayavka', 'text'],
              ['Дата поступления', 'datePost', 'date'],
              ['Количество мест', 'mesta', 'number'],
              ['Артикул (кол-во) (План)', 'artPlan', 'text'],
              ['Артикул (кол-во) (Факт)', 'artFact', 'text'],
              ['Единиц сырья (План)', 'syrPlan', 'number'],
              ['Единиц сырья (Факт)', 'syrFact', 'number'],
            ] as const).map(([label, k, type]) => (
              <div className="form-group" key={k}><label>{label}</label>
                <input type={type} min={type === 'number' ? 0 : undefined} value={modal.form[k]} onChange={setForm(k)}
                  placeholder={k === 'zayavka' ? 'Например: ДЛТ-Заявка от 15.06.2024' : undefined} /></div>
            ))}
            <div className="form-group"><label>Статус</label>
              <select value={modal.form.status} onChange={setForm('status')}>
                {STATUS_ORDER.map((s) => <option key={s} value={s}>{s}</option>)}
              </select></div>
            {([
              ['Дата передачи на производство', 'dateProd', 'date'],
              ['Дата готовности', 'dateReady', 'date'],
              ['Исполнитель', 'ispolnitel', 'text'],
            ] as const).map(([label, k, type]) => (
              <div className="form-group" key={k}><label>{label}</label>
                <input type={type} value={modal.form[k]} onChange={setForm(k)} /></div>
            ))}
            <div className="form-group"><label>Комментарий</label>
              <textarea value={modal.form.comment} onChange={setForm('comment')} /></div>
            <div className="form-group"><label>Заявки с отчета</label>
              <input type="text" value={modal.form.otchet} onChange={setForm('otchet')} /></div>
            <div className="form-group"><label>Идентификатор</label>
              <input type="text" value={modal.form.identifier} onChange={setForm('identifier')} /></div>
            <div className="modal-buttons">
              <button className="btn-cancel" onClick={() => setModal(null)}>Отмена</button>
              <button className="btn-save" onClick={saveModal}>Сохранить</button>
            </div>
          </div>
        </div>
      )}

      <div className={`toast ${toast.type} ${toast.show ? 'show' : ''}`}>{toast.msg}</div>
    </div>
  );
}
