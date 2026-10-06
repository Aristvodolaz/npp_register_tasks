import { BadRequestException, Injectable, NotFoundException } from '@nestjs/common';
import * as sql from 'mssql';
import { DbService } from '../db/db.service';

export const STATUSES = [
  'Новый', 'В работе на приемке', 'Передан на производство', 'В работе на производстве',
  'В работе у операторов', 'Выполнен', 'Отгружено', 'Заявка удалена',
];

type Kind = 'str' | 'int' | 'date' | 'status';
// поле DTO -> [колонка существующей таблицы dbo.npp_orders, тип, макс. длина]
const FIELDS: Record<string, [string, Kind, number]> = {
  zayavka: ['zayavka', 'str', 500],
  datePost: ['datePost', 'date', 20],
  mesta: ['mesta', 'int', 50],
  artPlan: ['artPlan', 'str', 500],
  artFact: ['artFact', 'str', 500],
  syrPlan: ['syrPlan', 'int', 50],
  syrFact: ['syrFact', 'int', 50],
  status: ['status', 'status', 100],
  dateProd: ['dateProd', 'date', 20],
  dateReady: ['dateReady', 'date', 20],
  ispolnitel: ['ispolnitel', 'str', 200],
  comment: ['comment', 'str', 0],
  customer: ['customer', 'str', 100],
  otchet: ['otchet', 'str', 500],
  identifier: ['order_id', 'str', 100],
  dateFromName: ['dateFromName', 'str', 50],
};

const SELECT = `SELECT db_id AS id, zayavka, datePost, mesta, artPlan, artFact, syrPlan, syrFact, status,
  dateProd, dateReady, ispolnitel, comment, customer, otchet, order_id AS identifier, dateFromName
  FROM dbo.npp_orders`;

/** Строка БД -> DTO (текстовые числа и даты приводятся к типам). */
function toDto(r: Record<string, any>) {
  return {
    ...r,
    datePost: toDate(r.datePost),
    dateProd: toDate(r.dateProd),
    dateReady: toDate(r.dateReady),
    mesta: toInt(r.mesta),
    syrPlan: toInt(r.syrPlan),
    syrFact: toInt(r.syrFact),
    zayavka: r.zayavka ?? '', artPlan: r.artPlan ?? '', artFact: r.artFact ?? '',
    ispolnitel: r.ispolnitel ?? '', comment: r.comment ?? '', otchet: r.otchet ?? '',
    identifier: r.identifier ?? '', customer: r.customer ?? '', dateFromName: r.dateFromName ?? '',
    status: r.status || 'Новый',
  };
}

const today = () => new Date().toISOString().slice(0, 10);

function toDate(v: unknown): string | null {
  if (v === null || v === undefined || v === '') return null;
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  let y: number, mo: number, d: number;
  if (m) { y = +m[1]; mo = +m[2]; d = +m[3]; }
  else if ((m = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/))) { d = +m[1]; mo = +m[2]; y = +m[3]; }
  else return null;
  const dt = new Date(Date.UTC(y, mo - 1, d));
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) return null;
  return dt.toISOString().slice(0, 10);
}

function toInt(v: unknown): number | null {
  if (v === null || v === undefined || v === '') return null;
  const n = Math.trunc(Number(v));
  return Number.isFinite(n) && Math.abs(n) < 2_000_000_000 ? n : null;
}

/** Оставляет только известные поля и приводит их к типам БД. */
function clean(input: Record<string, unknown>): Record<string, any> {
  const out: Record<string, any> = {};
  for (const [key, [, kind, max]] of Object.entries(FIELDS)) {
    if (!(key in input)) continue;
    const v = input[key];
    switch (kind) {
      case 'str':
        out[key] = max ? String(v ?? '').slice(0, max) : String(v ?? '');
        break;
      case 'int':
        out[key] = toInt(v);
        break;
      case 'date':
        out[key] = toDate(v);
        break;
      case 'status':
        if (!STATUSES.includes(String(v))) throw new BadRequestException(`Неизвестный статус: ${v}`);
        out[key] = String(v);
        break;
    }
  }
  return out;
}

/** Автоустановка дат при смене статуса (как в исходной версии). */
function autoDates(d: Record<string, any>, current?: Record<string, any>) {
  if (!d.status) return;
  const prod = 'dateProd' in d ? d.dateProd : current?.dateProd;
  const ready = 'dateReady' in d ? d.dateReady : current?.dateReady;
  if (d.status === 'Передан на производство' && !prod) d.dateProd = today();
  if ((d.status === 'Выполнен' || d.status === 'Заявка удалена') && !ready) d.dateReady = today();
}

function bind(req: sql.Request, key: string, v: any) {
  const max = FIELDS[key][2];
  const text = v === null || v === undefined ? null : String(v);
  return req.input(key, max ? sql.NVarChar(max) : sql.NVarChar(sql.MAX), text);
}

@Injectable()
export class OrdersService {
  constructor(private readonly db: DbService) {}

  async findAll() {
    return (await this.db.query(`${SELECT} ORDER BY db_id`)).map(toDto);
  }

  private async findOne(id: number) {
    const r = await this.db.request().input('id', sql.Int, id).query(`${SELECT} WHERE db_id = @id`);
    if (!r.recordset.length) throw new NotFoundException('Заявка не найдена');
    return toDto(r.recordset[0]);
  }

  async create(input: Record<string, unknown>) {
    const d = clean(input);
    d.status ??= 'Новый';
    autoDates(d);
    return this.findOne(await this.insert(d, this.db.request()));
  }

  async createMany(inputs: Record<string, unknown>[]) {
    const tx = new sql.Transaction(this.db.pool);
    await tx.begin();
    try {
      let imported = 0;
      for (const input of inputs) {
        const d = clean(input);
        if (!d.zayavka) continue;
        d.status ??= 'Новый';
        await this.insert(d, new sql.Request(tx));
        imported++;
      }
      await tx.commit();
      return { imported, skipped: inputs.length - imported };
    } catch (e) {
      await tx.rollback();
      throw e;
    }
  }

  private async insert(d: Record<string, any>, req: sql.Request): Promise<number> {
    const keys = Object.keys(d);
    keys.forEach((k) => bind(req, k, d[k]));
    const cols = keys.map((k) => FIELDS[k][0]).join(',');
    const vals = keys.map((k) => '@' + k).join(',');
    const r = await req.query(`INSERT INTO dbo.npp_orders (${cols}) OUTPUT INSERTED.db_id AS id VALUES (${vals})`);
    return r.recordset[0].id;
  }

  async update(id: number, input: Record<string, unknown>) {
    const d = clean(input);
    const current = await this.findOne(id);
    autoDates(d, current);
    const keys = Object.keys(d);
    if (!keys.length) return current;
    const req = this.db.request().input('id', sql.Int, id);
    keys.forEach((k) => bind(req, k, d[k]));
    const sets = keys.map((k) => `${FIELDS[k][0]} = @${k}`).join(', ');
    await req.query(`UPDATE dbo.npp_orders SET ${sets}, updated_at = GETDATE() WHERE db_id = @id`);
    return this.findOne(id);
  }

  async removeMany(ids: number[]) {
    const safe = ids.map(Number).filter((n) => Number.isInteger(n));
    if (!safe.length) return 0;
    const r = await this.db.request().query(`DELETE FROM dbo.npp_orders WHERE db_id IN (${safe.join(',')})`);
    return r.rowsAffected[0];
  }

  async clear() {
    const r = await this.db.request().query('DELETE FROM dbo.npp_orders');
    return r.rowsAffected[0];
  }
}
